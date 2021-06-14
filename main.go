package main

import (
	"errors"
	"fmt"
	"html/template"
	"io/ioutil"
	"net"
	"net/http"
	"os"
	"os/exec"
	"regexp"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

type TableStruct struct {
	Headers []string   // table headers, the field/column in the first row
	Datas   [][]string // data start from the second row to the end
}

var (
	staticPages = []string{"index.html", "import.html", "match.html"}
	commands    = map[string]string{
		"windows": "cmd /c start",
		"darwin":  "open",
		"linux":   "xdg-open",
	}

	MatchFields         = []string{"省", "市", "区"}
	MergeField          = "订单.*号"
	ExpressCompanyField = "快递名称"

	mutex     sync.Mutex
	tableData TableStruct

	hintResponseTemplate = `<html><body><h1>%v</h1><button type="button" onclick="window.history.back(-1);" class="btn  default">返回</button></body></html>`
)

const (
	localhost    = "0.0.0.0"
	port         = 18888
	dataFileName = "data.xlsx"
)

func getExpressCompany(td TableStruct) []string {
	m := make(map[string]bool)
	idx := -1
	for i, header := range td.Headers {
		if header == ExpressCompanyField {
			idx = i
			break
		}
	}
	if idx == -1 {
		return nil
	}
	for _, data := range td.Datas {
		m[data[idx]] = true
	}
	ret := make([]string, 0, len(m))
	for key := range m {
		ret = append(ret, key)
	}
	return ret
}

// opens the specified URL in the default browser of the user.
func open(uri string) error {
	run, ok := commands[runtime.GOOS]
	if !ok {
		return fmt.Errorf("don't know how to open things on %s platform", runtime.GOOS)
	}

	cmd := exec.Command(run, uri)
	return cmd.Start()
}

func SliceContain(str string, strs []string) bool {
	for _, s := range strs {
		if s == str {
			return true
		}
	}
	return false
}

func SliceEqual(s1, s2 []string) bool {
	if len(s1) != len(s2) {
		return false
	}
	for i := range s1 {
		if s1[i] != s2[i] {
			return false
		}
	}
	return true
}

func IndexHandler(writer http.ResponseWriter, req *http.Request) {
	fmt.Println("IndexHandler")
	fmt.Println(req.RequestURI)
	fmt.Println(req.Method)
	path := req.RequestURI
	if path[0] == '/' {
		path = path[1:]
	}
	fmt.Println("path: " + path)

	// only allow to get these three pages "index.html" "import.html" "match.html"
	if !SliceContain(path, staticPages) {
		writer.WriteHeader(http.StatusNotFound)
		return
	}

	t, err := template.New(path).Funcs(template.FuncMap{"getExpressCompany": getExpressCompany}).ParseFiles(path)
	if err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	// read previous data from the data file
	td, err := File2TableStruct(dataFileName)
	if err != nil {
		fmt.Println("read data from file failed. err: ", err)
	} else {
		mutex.Lock()
		tableData = td
		mutex.Unlock()
	}

	req.ParseForm()
	keyword := req.PostFormValue("keyword")
	fmt.Println("keyword: ", keyword)

	mutex.Lock()
	defer mutex.Unlock()
	res := tableData
	if keyword != "" {
		for i := 0; i < len(res.Datas); {
			matched := false
			for _, cell := range res.Datas[i] {
				m, err := regexp.MatchString(keyword, cell)
				if err == nil && m {
					matched = true
					break
				}
			}
			if !matched {
				res.Datas = append(res.Datas[:i], res.Datas[i+1:]...)
			} else {
				i++
			}
		}
	}
	if err := t.Execute(writer, res); err != nil {
		writer.Write([]byte(err.Error()))
		return
	}
}

func ImportHandler(writer http.ResponseWriter, req *http.Request) {
	fmt.Println("ImportHandler")
	fmt.Println(req.RequestURI)
	fmt.Println(req.Method)

	req.ParseForm()

	data1, err := getFormFileData(req, "upload_import")
	if err != nil && err != http.ErrMissingFile {
		writer.Write([]byte(err.Error()))
		return
	}

	data2, err := getFormFileData(req, "upload_import_2")
	if err != nil && err != http.ErrMissingFile {
		writer.Write([]byte(err.Error()))
		return
	}

	if data1 == nil && data2 == nil {
		writer.Write([]byte("Please upload at least one file"))
		return
	}

	if data1 != nil && data2 != nil {
		var td TableStruct
		td1, err := Byte2TableStruct(data1)
		if err != nil {
			writer.Write([]byte(err.Error()))
			return
		}
		td2, err := Byte2TableStruct(data2)
		if err != nil {
			writer.Write([]byte(err.Error()))
			return
		}

		td, err = match(td1, td2, []string{MergeField})
		if err != nil {
			writer.Write([]byte("Please upload at least one file"))
			return
		}

		matchIdx := make([]int, 0)
		for j, header := range td.Headers {
			m, err := regexp.MatchString(MergeField, header)
			if err == nil && m {
				matchIdx = append(matchIdx, j)
			}
		}

		td.Headers = filter(td.Headers, matchIdx)
		for i := range td.Datas {
			td.Datas[i] = filter(td.Datas[i], matchIdx)
		}

		data1, err = TableStruct2Byte(td)
		if err != nil {
			writer.Write([]byte(err.Error()))
			return
		}
	} else if data1 == nil {
		data1 = data2
	}

	handleImport(writer, req, data1)
}

func filter(data []string, idx []int) []string {
	length := 0
	for i := 0; i < len(data); i++ {
		found := false
		for j := 0; j < len(idx); j++ {
			if idx[j] == i {
				found = true
				break
			}
		}
		if !found {
			data[length] = data[i]
			length++
		}
	}
	return data[:length]
}

func getFormFileData(req *http.Request, key string) ([]byte, error) {
	file, _, err := req.FormFile(key)
	if err != nil {
		return nil, err
	}

	defer file.Close()

	data, err := ioutil.ReadAll(file)
	if err != nil {
		return nil, err
	}

	return data, nil
}

func MatchHandler(writer http.ResponseWriter, req *http.Request) {
	fmt.Println("MatchHandler")
	fmt.Println(req.RequestURI)
	fmt.Println(req.Method)

	req.ParseForm()
	data, err := getFormFileData(req, "upload_import")
	if err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	expressCompany := req.FormValue("expressComp")

	handleMatch(writer, req, data, expressCompany)
}

func ClearHandler(writer http.ResponseWriter, req *http.Request) {
	mutex.Lock()
	defer mutex.Unlock()
	tableData = TableStruct{}
	os.Remove(dataFileName)

	host := strings.Split(req.Host, ":")

	target := fmt.Sprintf("http://%v:%v/index.html", host[0], port)
	fmt.Printf("ClearHandler: %v\n", target)

	http.Redirect(writer, req, target, http.StatusTemporaryRedirect)
}

func DeleteHandler(writer http.ResponseWriter, req *http.Request) {
	defer func() {
		host := strings.Split(req.Host, ":")

		target := fmt.Sprintf("http://%v:%v/index.html", host[0], port)
		fmt.Printf("DeleteHandler: %v\n", target)

		http.Redirect(writer, req, target, http.StatusTemporaryRedirect)
	}()

	req.ParseForm()
	idStrArr := req.Form["id"]

	var ids []int
	for _, ele := range idStrArr {
		id, err := strconv.ParseInt(ele, 10, 64)
		if err != nil {
			fmt.Printf("strconv int failed. err:%v, ele:%v\n", err, ele)
			return
		}
		ids = append(ids, int(id))
	}

	sort.Ints(ids)

	mutex.Lock()
	defer mutex.Unlock()

	for i := len(ids) - 1; i >= 0; i-- {
		id := ids[i]
		if id >= len(tableData.Datas) {
			continue
		}
		tableData.Datas = append(tableData.Datas[:id], tableData.Datas[id+1:]...)
	}

	TableStruct2File(tableData, dataFileName) // ignore error
}

func handleMatch(writer http.ResponseWriter, req *http.Request, data []byte, express string) {
	mutex.Lock()
	tmpData := tableData
	mutex.Unlock()

	if len(tmpData.Headers) == 0 || len(tmpData.Datas) == 0 {
		hint := fmt.Sprintf(hintResponseTemplate, "请先上传数据")
		writer.Write([]byte(hint))
		return
	}

	// get input execel file and parse it from form data
	td, err := Byte2TableStruct(data)
	if err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	// match the input excel file with the imported data, and return the matched part data
	td, err = match(filterExpressCompany(tmpData, express), td, MatchFields)
	if err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	// transfer the result to excel file to return to front end
	xlsxFile := TableStruct2Excel(td)

	writer.Header().Set("Content-Type", "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	writer.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"", "匹配数据"+time.Now().Format("20060102150405")+".xlsx"))
	Excel2Writer(xlsxFile, writer)
}

func filterExpressCompany(data TableStruct, express string) TableStruct {
	if express == "" {
		return data
	}

	var ret TableStruct
	ret.Headers = data.Headers
	idx := -1
	for i, h := range ret.Headers {
		if h == ExpressCompanyField {
			idx = i
			break
		}
	}
	if idx == -1 {
		return data
	}
	for _, d := range data.Datas {
		if d[idx] == express {
			ret.Datas = append(ret.Datas, d)
		}
	}
	return ret
}

func match(data TableStruct, match TableStruct, matchFields []string) (TableStruct, error) {
	ret := match

	matchIdx := make([][]int, len(matchFields))
	for i, matchField := range matchFields {
		matchIdx[i] = nil
		for j, header := range data.Headers {
			m, err := regexp.MatchString(matchField, header)
			if err == nil && m {
				matchIdx[i] = []int{j, -1}
				break
			}
		}
		if matchIdx[i] == nil {
			return ret, errors.New("导入的数据无\"" + matchField + "\"字段")
		}

		for j, header := range match.Headers {
			m, err := regexp.MatchString(matchField, header)
			if err == nil && m {
				matchIdx[i][1] = j
				break
			}
		}
		if matchIdx[i][1] == -1 {
			return ret, errors.New("导入的匹配数据无\"" + matchField + "\"字段")
		}
	}

	ret.Headers = append(ret.Headers, data.Headers...)

	usedMap := make(map[int]bool)
	haveMathcedRow := make(map[int]bool)
	for i, datas := range match.Datas {
		for j, row := range data.Datas {
			if usedMap[j] {
				continue
			}

			match := true
			for _, idx := range matchIdx {
				if row[idx[0]] != datas[idx[1]] || row[idx[0]] == "" {
					match = false
					break
				}
			}
			if match {
				usedMap[j] = true
				ret.Datas[i] = append(ret.Datas[i], row...)
				haveMathcedRow[i] = true
				break
			}
		}
	}

	if len(matchFields) == 1 {
		return ret, nil
	}

	// search for second time to match remainder row, ignore the last match filed
	matchIdx = matchIdx[:len(matchIdx)-1]
	for i, datas := range match.Datas {
		if haveMathcedRow[i] {
			continue
		}
		for j, row := range data.Datas {
			if usedMap[j] {
				continue
			}

			match := true
			for _, idx := range matchIdx {
				if row[idx[0]] != datas[idx[1]] || row[idx[0]] == "" {
					match = false
					break
				}
			}
			if match {
				usedMap[j] = true
				ret.Datas[i] = append(ret.Datas[i], row...)
				haveMathcedRow[i] = true
				break
			}
		}
	}

	return ret, nil
}

func handleImport(writer http.ResponseWriter, req *http.Request, data []byte) {
	// get input execel file and parse it from form data
	td, err := Byte2TableStruct(data)
	if err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	// filter NULL or empty row
	size := 0
	for i := 0; i < len(td.Datas); i++ {
		notEmpty := true
		for j := 0; j < len(td.Datas[i]); j++ {
			str := strings.TrimSpace(td.Datas[i][j])
			if str == "" || strings.ToLower(str) == "null" {
				notEmpty = false
				break
			}
		}
		if notEmpty {
			td.Datas[size] = td.Datas[i]
			size++
		}
	}
	td.Datas = td.Datas[:size]

	mutex.Lock()
	defer mutex.Unlock()

	if tableData.Headers == nil || tableData.Datas == nil {
		tableData = td
	} else {
		td, err = alignTableData(tableData, td)
		if err != nil {
			writer.Write([]byte(err.Error()))
			return
		}

		// compare the table header with the imported data, we only append the imported data to the
		// current data if the table header is the same
		if !SliceEqual(tableData.Headers, td.Headers) {
			respData := fmt.Sprintf(hintResponseTemplate, "表头不匹配，请先清空数据库或上传相同表头的数据")
			writer.Write([]byte(respData))
			return
		}

		for _, row := range td.Datas {
			isDiff := true
			for _, r := range tableData.Datas {
				if SliceEqual(r, row) {
					isDiff = false
					break
				}
			}
			if isDiff {
				tableData.Datas = append(tableData.Datas, row)
			}
		}
	}

	// open the file named "data.txt" to store the import data
	if err := TableStruct2File(tableData, dataFileName); err != nil {
		writer.Write([]byte(err.Error()))
		return
	}

	writer.Write([]byte(fmt.Sprintf(hintResponseTemplate, "上传数据成功")))
}

// 调整 td 表头的顺序使其于 origin 的表头顺序相同
func alignTableData(origin TableStruct, td TableStruct) (TableStruct, error) {
	if len(origin.Headers) != len(td.Headers) {
		return td, nil
	}

	alignIdxs := make([]int, len(origin.Headers))
	for i := range origin.Headers {
		alignIdxs[i] = -1
		for j := range td.Headers {
			if td.Headers[j] == origin.Headers[i] {
				alignIdxs[i] = j
				break
			}
		}
		if alignIdxs[i] == -1 {
			return td, fmt.Errorf(hintResponseTemplate, "表头不匹配，请先清空数据库或上传相同表头的数据")
		}
	}

	var ret TableStruct
	ret.Headers = append(ret.Headers, origin.Headers...)
	for i := range td.Datas {
		data := make([]string, len(td.Datas[i]))
		for j, idx := range alignIdxs {
			data[j] = td.Datas[i][idx]
		}
		ret.Datas = append(ret.Datas, data)
	}

	return ret, nil
}

func main() {
	mux := http.NewServeMux()
	mux.HandleFunc("/api/import/", ImportHandler)
	mux.HandleFunc("/api/match/", MatchHandler)
	mux.HandleFunc("/api/clear/", ClearHandler)
	mux.HandleFunc("/api/delete", DeleteHandler)

	http.HandleFunc("/api/", mux.ServeHTTP)
	http.HandleFunc("/index.html", IndexHandler)
	http.HandleFunc("/match.html", IndexHandler)
	http.HandleFunc("/import.html", IndexHandler)

	addr := localhost + ":" + strconv.FormatInt(port, 10)

	go func() {
		for i := 0; i < 5; i++ {
			tcpAddr, _ := net.ResolveTCPAddr("tcp", addr)
			tcpConn, err := net.DialTCP("tcp", nil, tcpAddr)
			if err != nil {
				fmt.Printf("i:%v, err: %v\n", i, err)
				continue
			}
			tcpConn.Close()
			break
		}
		open("http://" + localhost + ":" + strconv.FormatInt(port, 10) + "/index.html")
	}()

	http.ListenAndServe(addr, nil)
}
