package main

import (
	"embed"
	"fmt"
	"io/fs"
	"log"
	"os/user"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/net/context"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

var (
	user32                       = syscall.NewLazyDLL("user32.dll")
	kernel32                     = syscall.NewLazyDLL("kernel32.dll")
	psapi                        = syscall.NewLazyDLL("psapi.dll")
	version                      = syscall.NewLazyDLL("version.dll")
	procGetForegroundWindow      = user32.NewProc("GetForegroundWindow")
	procGetWindowTextW           = user32.NewProc("GetWindowTextW")
	procGetWindowThreadProcessId = user32.NewProc("GetWindowThreadProcessId")
	procOpenProcess              = kernel32.NewProc("OpenProcess")
	procGetModuleFileNameEx      = psapi.NewProc("GetModuleFileNameExW")
	procGetFileVersionInfoSize   = version.NewProc("GetFileVersionInfoSizeW")
	procGetFileVersionInfo       = version.NewProc("GetFileVersionInfoW")
	procVerQueryValue            = version.NewProc("VerQueryValueW")
	PROCESS_QUERY_INFORMATION    = 0x0400
	PROCESS_VM_READ              = 0x0010
)

func getForegroundWindow() syscall.Handle {
	ret, _, _ := procGetForegroundWindow.Call()
	return syscall.Handle(ret)
}

func getWindowText(hwnd syscall.Handle) string {
	buf := make([]uint16, 256) // Buffer size for the window title
	ret, _, _ := procGetWindowTextW.Call(
		uintptr(hwnd),
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(len(buf)),
	)
	if ret == 0 {
		return ""
	}
	return syscall.UTF16ToString(buf)
}

func getProcessID(hwnd syscall.Handle) uint32 {
	var pid uint32
	procGetWindowThreadProcessId.Call(
		uintptr(hwnd),
		uintptr(unsafe.Pointer(&pid)),
	)
	return pid
}

func getProcessName(pid uint32) string {
	handle, _, _ := procOpenProcess.Call(
		uintptr(PROCESS_QUERY_INFORMATION|PROCESS_VM_READ),
		uintptr(0),
		uintptr(pid),
	)
	if handle == 0 {
		return ""
	}
	defer syscall.CloseHandle(syscall.Handle(handle))

	var buf [1024]uint16
	ret, _, _ := procGetModuleFileNameEx.Call(
		handle,
		0,
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(1024),
	)
	if ret == 0 {
		return ""
	}
	exePath := syscall.UTF16ToString(buf[:])

	// Get version information size
	size, _, _ := procGetFileVersionInfoSize.Call(
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(exePath))),
		0,
	)
	if size == 0 {
		return ""
	}

	// Allocate buffer for version information
	verInfo := make([]byte, size)
	ret, _, _ = procGetFileVersionInfo.Call(
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(exePath))),
		0,
		size,
		uintptr(unsafe.Pointer(&verInfo[0])),
	)
	if ret == 0 {
		return ""
	}

	// Query value to get FileDescription
	var bufferLen uint32
	var buffer uintptr
	ret, _, _ = procVerQueryValue.Call(
		uintptr(unsafe.Pointer(&verInfo[0])),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(`\StringFileInfo\040904b0\FileDescription`))),
		uintptr(unsafe.Pointer(&buffer)),
		uintptr(unsafe.Pointer(&bufferLen)),
	)
	if ret == 0 {
		return ""
	}
	description := (*[1 << 20]uint16)(unsafe.Pointer(buffer))[:bufferLen:bufferLen]
	return syscall.UTF16ToString(description)
}

func getProcessPath(pid uint32) string {
	handle, _, _ := procOpenProcess.Call(
		uintptr(PROCESS_QUERY_INFORMATION|PROCESS_VM_READ),
		uintptr(0),
		uintptr(pid),
	)
	if handle == 0 {
		return ""
	}
	defer syscall.CloseHandle(syscall.Handle(handle))

	var buf [1024]uint16
	ret, _, _ := procGetModuleFileNameEx.Call(
		handle,
		0,
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(len(buf)),
	)
	if ret == 0 {
		return ""
	}
	return syscall.UTF16ToString(buf[:])
}

func processTitle(processName, title string) (string, string) {
	if processName == "" {
		if idx := strings.Index(title, "-"); idx != -1 {
			// Split the title at the hyphen
			processName = strings.TrimSpace(title[:idx]) // Everything before the hyphen
			title = strings.TrimSpace(title[idx+1:])     // Everything after the hyphen
		} else {
			// If no hyphen is found, set the process name to the title
			processName = title
		}
	}
	return processName, title
}

//go:embed credentials.json
var credentialsFS embed.FS

func main() {

	fmt.Print("Service started\n")
	currentUser, err := user.Current()
	if err != nil {
		log.Fatalf("Failed to get current user: %v", err)
	}

	ctx := context.Background()

	// Access the embedded credentials.json file
	b, err := fs.ReadFile(credentialsFS, "credentials.json")
	if err != nil {
		log.Fatalf("Failed to read embedded credentials: %v", err)
	}

	// Assuming the credentials are for a service account or a previously saved OAuth token
	config, err := google.JWTConfigFromJSON(b, sheets.SpreadsheetsScope)
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}

	//client := getClient(config)
	client := config.Client(ctx) // Now using 'ctx'

	// Create a new Sheets service.
	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	// Define the ID of the spreadsheet and the range you want to write to
	spreadsheetId := "1_kwLZEBi9xh_IAZ3lEKeHSRlOprrkDIwt8BI8iuZt9Q"
	writeRange := "Sheet1!A2" // Example range, adjust as needed
	// How the input data should be interpreted.
	valueInputOption := "USER_ENTERED"

	// How the input data should be inserted.
	insertDataOption := "INSERT_ROWS"

	var lastTitle string
	for {
		hwnd := getForegroundWindow()
		title := getWindowText(hwnd)
		if title != lastTitle {
			lastTitle = title
			pid := getProcessID(hwnd)
			processName := getProcessName(pid)

			processName, title = processTitle(processName, title)

			fmt.Printf("New active app: %s\n", processName)
			fmt.Printf("New active window: %s\n", title)

			currentTime := time.Now()
			formattedTime := currentTime.Format("1/2/2006 15:04:05")

			var vr sheets.ValueRange
			myval := []interface{}{formattedTime, currentUser.Name, processName, title}
			vr.Values = append(vr.Values, myval)

			_, err = srv.Spreadsheets.Values.Append(spreadsheetId, writeRange, &vr).ValueInputOption(valueInputOption).InsertDataOption(insertDataOption).Do()
			if err != nil {
				log.Fatalf("Unable to write data to sheet: %v", err)
			}

		}
		time.Sleep(1 * time.Second) // Check every second
	}
}
