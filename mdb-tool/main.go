package main

import (
	"encoding/json"
	"fmt"
	"github.com/LeeZXin/zsf/util/threadutil"
	_ "github.com/alexbrainman/odbc"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	_ "github.com/mattn/go-adodb"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"runtime"
	"strconv"
	"time"
)

type searchData struct {
	TableName string           `json:"tableName"`
	Data      []map[string]any `json:"data"`
	VarList   []string         `json:"varList"`
}

func searchProcess(filePath string, varName string, val string, fn ...func([]searchData, error)) error {
	return threadutil.RunSafe(func() {
		err := operateAccess(filePath, func(db *ole.IDispatch) {
			tableNames, err := getAllTablesNames(db)
			if err != nil {
				if fn != nil {
					for _, f := range fn {
						f(nil, err)
					}
				}
				return
			}
			ret := make([]searchData, 0)
			for _, name := range tableNames {
				data, vars, err := search(db, name, varName, val)
				if err == nil {
					if len(data) > 0 {
						ret = append(ret, searchData{
							TableName: name,
							Data:      data,
							VarList:   vars,
						})
					}
				}
			}
			if fn != nil {
				for _, f := range fn {
					f(ret, nil)
				}
			}
		})
		if err != nil {
			os.Exit(1)
		}
	})
}

func isSystemTable(tableName string) bool {
	return tableName[:4] == "MSys" || tableName[:4] == "~TMP"
}

func isAbnormalTable(tableName string) bool {
	return tableName[:1] == "~" || tableName[:1] == "_"
}

func operateAccess(path string, fn ...func(*ole.IDispatch)) error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	// 初始化 COM 库
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	// 创建 Access 应用程序对象
	unknown, err := oleutil.CreateObject("DAO.DBEngine.120")
	if err != nil {
		unknown, err = oleutil.CreateObject("DAO.DBEngine.36")
		if err != nil {
			unknown, err = oleutil.CreateObject("DAO.DBEngine.35")
			if err != nil {
				unknown, err = oleutil.CreateObject("DAO.DBEngine.30")
				if err != nil {
					unknown, err = oleutil.CreateObject("DAO.DBEngine.12")
					if err != nil {
						fmt.Println("DAO.DBEngine not correct")
						return err
					}
				}
			}
		}
	}
	access, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer access.Release()
	path, err = filepath.Abs(path)
	if err != nil {
		return err
	}
	// 打开Access数据库
	db, err := oleutil.CallMethod(access, "OpenDatabase", path)
	if err != nil {
		return err
	}
	if fn != nil {
		for _, f := range fn {
			f(db.ToIDispatch())
		}
	}
	return nil
}

func search(db *ole.IDispatch, tableName string, varName string, value any) ([]map[string]any, []string, error) {
	var sql string
	if varName == "" {
		sql = fmt.Sprintf("select * from %s", tableName)
	} else {
		sql = fmt.Sprintf("select * from %s where %s = '%v'", tableName, varName, value)
	}
	return querySql(db, sql)
}

func querySql(db *ole.IDispatch, sql string) ([]map[string]any, []string, error) {
	query, err := oleutil.CallMethod(db, "OpenRecordset", sql)
	if err != nil {
		return nil, nil, err
	}
	defer query.Clear()
	ret := make([]map[string]any, 0)
	var varList []string
	var count int
	for {
		data := make(map[string]any, 8)
		queryIDispatch := query.ToIDispatch()
		eof, err := oleutil.GetProperty(queryIDispatch, "EOF")
		if err != nil {
			return nil, nil, err
		}
		if eof.Value().(bool) {
			break
		}
		fields, err := oleutil.GetProperty(queryIDispatch, "Fields")
		if err != nil {
			return nil, nil, err
		}
		fieldsIDispatch := fields.ToIDispatch()
		if varList == nil {
			fieldsCount, err := oleutil.GetProperty(fieldsIDispatch, "Count")
			if err != nil {
				return nil, nil, err
			}
			fieldsCount.Clear()
			count = int(fieldsCount.Val)
			varList = make([]string, count)
			for i := 0; i < count; i++ {
				item, err := oleutil.GetProperty(fieldsIDispatch, "Item", i)
				if err != nil {
					return nil, nil, err
				}
				name, err := oleutil.GetProperty(item.ToIDispatch(), "Name")
				if err != nil {
					return nil, nil, err
				}
				varList[i] = name.ToString()
				item.Clear()
				name.Clear()
			}
		}
		for i := 0; i < count; i++ {
			item, err := oleutil.GetProperty(queryIDispatch, "Fields", i)
			if err != nil {
				return nil, nil, err
			}
			itemIDispatch := item.ToIDispatch()
			value, err := oleutil.GetProperty(itemIDispatch, "Value")
			if err != nil {
				return nil, nil, err
			}
			data[varList[i]] = value.Value()
			item.Clear()
			value.Clear()
		}
		ret = append(ret, data)
		fields.Clear()
		queryIDispatch.CallMethod("MoveNext")
	}
	return ret, varList, nil
}

func getAllTablesNames(db *ole.IDispatch) ([]string, error) {
	// 获取表的集合对象
	tables, err := oleutil.GetProperty(db, "TableDefs")
	if err != nil {
		return nil, err
	}
	// 获取表的数量
	count, err := oleutil.GetProperty(tables.ToIDispatch(), "Count")
	if err != nil {
		return nil, err
	}
	tableCount := int(count.Val)
	ret := make([]string, 0)
	for i := 0; i < tableCount; i++ {
		table, err := oleutil.GetProperty(tables.ToIDispatch(), "Item", i)
		if err != nil {
			return nil, err
		}
		tableName, err := oleutil.GetProperty(table.ToIDispatch(), "Name")
		if err != nil {
			return nil, err
		}
		tableNameStr := tableName.Value().(string)
		if !isSystemTable(tableNameStr) && !isAbnormalTable(tableNameStr) {
			ret = append(ret, tableNameStr)
		}
		table.Clear()
		tableName.Clear()
	}
	tables.Clear()
	return ret, nil
}

func saveToExcel(dataList []searchData, dirName string) {
	f := excelize.NewFile()
	defer f.Close()
	// Create a new sheet.
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		return
	}
	f.SetActiveSheet(index)
	row := 1
	for _, data := range dataList {
		/*f.SetCellValue("Sheet1", "A"+strconv.Itoa(row), data.TableName)
		row += 1
		for i, v := range data.VarList {
			name, _ := excelize.ColumnNumberToName(i)
			f.SetCellValue("Sheet1", name+strconv.Itoa(row), v)
		}
		row += 1*/
		for _, d := range data.Data {
			for i, v := range data.VarList {
				name, _ := excelize.ColumnNumberToName(i)
				f.SetCellValue("Sheet1", name+strconv.Itoa(row), d[v])
			}
			row += 1
		}
	}
	// Save spreadsheet by the given path.
	fileName := dirName + "\\" + time.Now().Format("20060102150405") + "_查询数据.xlsx"
	f.SaveAs(fileName)
}

func cmdProcess(filePath string, vars []string) {
	cmd := vars[0]
	switch cmd {
	case "excel":
		searchProcess(filePath, vars[1], vars[2], func(data []searchData, err error) {
			if err != nil {
				return
			}
			saveToExcel(data, vars[3])
		})
	case "search":
		searchProcess(filePath, vars[1], vars[2], func(data []searchData, err error) {
			if err != nil {
				fmt.Println(err)
				return
			}
			m, _ := json.Marshal(data)
			os.WriteFile("search.json", m, os.ModePerm)
		})
	}
}

func main() {
	cmdProcess(os.Args[1], os.Args[2:])
}
