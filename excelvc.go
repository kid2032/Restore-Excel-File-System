package main

import (
	"bytes"
	"compress/gzip"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"crypto/sha256"
	"database/sql"

	"encoding/hex"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/fsnotify/fsnotify"
	"github.com/getlantern/systray"
	_ "github.com/go-sql-driver/mysql"
)

var (
	db *sql.DB

	mu       sync.Mutex
	debounce = make(map[string]*time.Timer)
	watchMu  sync.Mutex
	watchers = make(map[string]*fsnotify.Watcher)
)

/* =================== MAIN =================== */

func main() {
	logFile, _ := os.OpenFile("excelvc.log",
		os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	log.SetOutput(logFile)

	var err error
	db, err = sql.Open("mysql",
		"root:root123@tcp(localhost:3306)/excelvc?parseTime=true")
	if err != nil {
		log.Fatal(err)
	}

	go startCleanup()

	go startWeb()
	systray.Run(onTrayReady, func() {})
}

/* =================== TRAY =================== */

func onTrayReady() {
	systray.SetTitle("ExcelVC")
	systray.SetTooltip("Excel Version Control")

	open := systray.AddMenuItem("Open Web UI", "")
	quit := systray.AddMenuItem("Quit", "")

	go func() {
		for {
			select {
			case <-open.ClickedCh:
				openBrowser("http://localhost:8080")
			case <-quit.ClickedCh:
				os.Exit(0)
			}
		}
	}()
}

/* =================== WEB =================== */

func startWeb() {
	http.HandleFunc("/", home)
	http.HandleFunc("/watch", watchHandler)
	http.HandleFunc("/versions", listVersions)
	http.HandleFunc("/restore", restoreVersion)

	log.Println("Running at http://localhost:8080")
	http.ListenAndServe(":8080", nil)
}
func encrypt(data []byte) []byte {
	key := []byte(os.Getenv("EXCELVC_KEY"))

	block, _ := aes.NewCipher(key)
	gcm, _ := cipher.NewGCM(block)

	nonce := make([]byte, gcm.NonceSize())
	rand.Read(nonce)

	return gcm.Seal(nonce, nonce, data, nil)
}

func decrypt(data []byte) []byte {
	key := []byte(os.Getenv("EXCELVC_KEY"))

	block, _ := aes.NewCipher(key)
	gcm, _ := cipher.NewGCM(block)

	nonceSize := gcm.NonceSize()
	nonce, cipherText := data[:nonceSize], data[nonceSize:]

	out, _ := gcm.Open(nil, nonce, cipherText, nil)
	return out
}

func home(w http.ResponseWriter, _ *http.Request) {
	fmt.Fprintln(w, `<html><body>
	<h2>Excel Version Control</h2>

	<form method="POST" action="/watch">
		<input name="path" style="width:400px" placeholder="C:\MyFolder">
		<button>Watch Folder</button>
	</form><hr>`)

	rows, _ := db.Query("SELECT id, file_name FROM excel_files")
	for rows.Next() {
		var id int
		var name string
		rows.Scan(&id, &name)
		fmt.Fprintf(w, `<a href="/versions?id=%d">%s</a><br>`, id, name)
	}
	fmt.Fprintln(w, "</body></html>")
}

/* =================== WATCH =================== */
func listVersions(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	if id == "" {
		fmt.Fprintln(w, "Missing file id")
		return
	}

	rows, err := db.Query(`
		SELECT version_number, created_at
		FROM excel_versions
		WHERE file_id=?
		ORDER BY version_number DESC
	`, id)
	if err != nil {
		fmt.Fprintln(w, "Database error")
		return
	}
	defer rows.Close()

	fmt.Fprintln(w, "<html><body>")
	fmt.Fprintln(w, "<h2>Versions</h2>")

	for rows.Next() {
		var v int
		var t time.Time
		rows.Scan(&v, &t)

		fmt.Fprintf(
			w,
			`<a href="/restore?id=%s&v=%d">Restore version %d (%s)</a><br>`,
			id,
			v,
			v,
			t.Format("2006-01-02 15:04:05"),
		)
	}

	fmt.Fprintln(w, `<br><a href="/">Back</a>`)
	fmt.Fprintln(w, "</body></html>")
}

func watchHandler(w http.ResponseWriter, r *http.Request) {
	path := r.FormValue("path")

	info, err := os.Stat(path)
	if err != nil || !info.IsDir() {
		fmt.Fprintln(w, "Invalid folder")
		return
	}

	watchMu.Lock()
	if _, ok := watchers[path]; ok {
		watchMu.Unlock()
		fmt.Fprintln(w, "Already watching this folder")
		return
	}
	watchMu.Unlock()

	go watchRecursive(path)
	fmt.Fprintln(w, "Now watching:", path)
}

func watchRecursive(root string) {
	watcher, _ := fsnotify.NewWatcher()

	filepath.WalkDir(root, func(p string, d os.DirEntry, _ error) error {
		if d != nil && d.IsDir() {
			watcher.Add(p)
		}
		return nil
	})

	watchMu.Lock()
	watchers[root] = watcher
	watchMu.Unlock()

	for {
		select {
		case e := <-watcher.Events:
			handleEvent(e)
		}
	}
}

func handleEvent(e fsnotify.Event) {
	if !isExcel(e.Name) || strings.HasPrefix(filepath.Base(e.Name), "~$") {
		return
	}

	mu.Lock()
	if t, ok := debounce[e.Name]; ok {
		t.Stop()
	}
	debounce[e.Name] = time.AfterFunc(1*time.Second, func() {
		saveVersion(e.Name)
	})
	mu.Unlock()
}

/* =================== SAVE =================== */

func saveVersion(path string) {
	waitStable(path)

	data, err := os.ReadFile(path)
	if err != nil {
		return
	}

	hash := sha(data)

	mu.Lock()
	defer mu.Unlock()

	tx, _ := db.Begin()
	defer tx.Rollback()

	var id int
	var lastHash sql.NullString

	err = tx.QueryRow(
		"SELECT id, last_hash FROM excel_files WHERE file_path=?",
		path).Scan(&id, &lastHash)

	if err == sql.ErrNoRows {
		r, _ := tx.Exec(
			"INSERT INTO excel_files(file_path,file_name,last_hash) VALUES(?,?,?)",
			path, filepath.Base(path), hash)
		i, _ := r.LastInsertId()
		id = int(i)
	} else if lastHash.Valid && lastHash.String == hash {
		return
	}

	var v int
	tx.QueryRow(
		"SELECT COALESCE(MAX(version_number),0)+1 FROM excel_versions WHERE file_id=?",
		id).Scan(&v)

	tx.Exec(
		"INSERT INTO excel_versions(file_id,version_number,file_data) VALUES(?,?,?)",
		id, v, encrypt(compress(data)),
	)

	tx.Exec(
		"UPDATE excel_files SET last_hash=? WHERE id=?", hash, id)

	tx.Commit()
}

/* =================== RESTORE =================== */

func restoreVersion(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	v := r.URL.Query().Get("v")

	var path string
	db.QueryRow("SELECT file_path FROM excel_files WHERE id=?", id).Scan(&path)

	if isLocked(path) {
		fmt.Fprintln(w, "File is open in Excel. Close it first.")
		return
	}

	var data []byte
	db.QueryRow(
		"SELECT file_data FROM excel_versions WHERE file_id=? AND version_number=?",
		id, v).Scan(&data)

	os.WriteFile(path, decompress(decrypt(data)), 0644)
	fmt.Fprintln(w, "Restored successfully. Reopen Excel.")
}

/* =================== HELPERS =================== */

func isExcel(p string) bool {
	ext := strings.ToLower(filepath.Ext(p))
	return ext == ".xls" || ext == ".xlsx"
}

func isLocked(path string) bool {
	f, err := os.OpenFile(path, os.O_WRONLY, 0666)
	if err != nil {
		return true
	}
	f.Close()
	return false
}

func waitStable(path string) {
	var last int64 = -1
	for {
		i, _ := os.Stat(path)
		if i.Size() == last {
			return
		}
		last = i.Size()
		time.Sleep(400 * time.Millisecond)
	}
}

func sha(b []byte) string {
	s := sha256.Sum256(b)
	return hex.EncodeToString(s[:])
}

func compress(b []byte) []byte {
	var buf bytes.Buffer
	g := gzip.NewWriter(&buf)
	g.Write(b)
	g.Close()
	return buf.Bytes()
}

func decompress(b []byte) []byte {
	r, _ := gzip.NewReader(bytes.NewReader(b))
	out, _ := io.ReadAll(r)
	r.Close()
	return out
}

func openBrowser(url string) {
	exec := "cmd"
	args := []string{"/c", "start", url}
	_, _ = os.StartProcess(exec, append([]string{exec}, args...), &os.ProcAttr{})
}
func startCleanup() {
	ticker := time.NewTicker(24 * time.Hour)
	for {
		<-ticker.C
		cleanupOldVersions()
	}
}

func cleanupOldVersions() {
	res, err := db.Exec(`
		DELETE FROM excel_versions
		WHERE created_at < NOW() - INTERVAL 7 DAY
	`)
	if err != nil {
		log.Println("Cleanup error:", err)
		return
	}

	n, _ := res.RowsAffected()
	log.Println("Cleanup removed versions:", n)
}
