// GreatShaman project main.go
package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

type MyMainWindow struct {
	*walk.MainWindow
	edit *walk.TextEdit
	path string
}

var edit *walk.TextEdit
var OOO map[string][]string
var patchfile = ""
var fname = ""

func main() {

	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println("405")
	}

	mw := &MyMainWindow{}

	MW := MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "GreatShaman",
		MinSize:  Size{500, 400},

		Size:   Size{500, 400},
		Layout: VBox{},
		Children: []Widget{
			TextEdit{
				MinSize:  Size{300, 300},
				MaxSize:  Size{1000, 1000},
				AssignTo: &mw.edit, ReadOnly: true,
			},
			PushButton{
				Text: "Вывести инструкцию",
				OnClicked: func() {
					s := fmt.Sprintf("Для создания документов администратору, откройте файл отгрузки и выберите желаемое действие: \r\n") + fmt.Sprintf("1:Admin: Создаст файлы администратора. \r\n") + fmt.Sprintf("2:Admin and Clients: Создаст файлы администратора и клиентов. \r\n") + fmt.Sprintf("3:Clients: Создаст файлы клиентов. \r\n\r\n") + fmt.Sprintf("Для создания ТТН, используйте кнопку(ТТН) и выберите нужный файл из папки Administrator. \r\n") + fmt.Sprintf("Это создаст выбранные вами документы в соответствующих папках. \r\n")
					mw.edit.SetText(s)

				},
			},
			PushButton{
				Text:      "Выбрать файл отгрузки",
				OnClicked: mw.pbClicked,
			},

			PushButton{
				Text: "Открыть папку с Программой.",
				OnClicked: func() {
					err = exec.Command("explorer", dir).Start()
				},
			},
			HSplitter{
				MinSize: Size{30, 20},
				MaxSize: Size{2000, 20},
				Children: []Widget{PushButton{
					Text: "Admin",
					OnClicked: func() {
						s := fmt.Sprintf("Запущено создание документов в папке Administrator.\r\nОжидайте пока это сообщение изменится что бы продолжить работу.")
						mw.edit.SetText(s)

						if len(patchfile) > 0 {
							var wg sync.WaitGroup
							wg.Add(1)
							// закрываем в анонимной функции переменную из цикла,что бы предотвартить её потерю во время обработки
							time.Sleep(15 * time.Millisecond)
							go func(patch string, flag int) {
								defer wg.Done()
								CreateMaster(patchfile, 1)
								mw.edit.SetText("Процедура завершена, вы можете продолжить работу.")
							}(patchfile, 1)

						} else {
							s := fmt.Sprintf("Сначала откройте подходящий файл отгрузки")
							mw.edit.SetText(s)
						}

					},
				}, PushButton{
					Text: "Admin And Clients",
					OnClicked: func() {
						s := fmt.Sprintf("Запущено создание документов в папках Admin и Clients.\r\nОжидайте пока это сообщение изменится что бы продолжить работу.")
						mw.edit.SetText(s)
						if len(patchfile) > 0 {

							var wg sync.WaitGroup
							wg.Add(1)
							// закрываем в анонимной функции переменную из цикла,что бы предотвартить её потерю во время обработки
							time.Sleep(15 * time.Millisecond)
							go func(patch string, flag int) {
								defer wg.Done()
								CreateMaster(patchfile, 3)
								mw.edit.SetText("Процедура завершена, вы можете продолжить работу.")
							}(patchfile, 3)
						} else {
							s := fmt.Sprintf("Сначала откройте подходящий файл отгрузки")
							mw.edit.SetText(s)
						}

					},
				}, PushButton{
					Text: "Clients",
					OnClicked: func() {
						s := fmt.Sprintf("Запущено создание документов в папке Clients.\r\nОжидайте пока это сообщение изменится что бы продолжить работу.")
						mw.edit.SetText(s)
						if len(patchfile) > 0 {

							var wg sync.WaitGroup
							wg.Add(1)
							// закрываем в анонимной функции переменную из цикла,что бы предотвартить её потерю во время обработки
							time.Sleep(15 * time.Millisecond)
							go func(patch string, flag int) {
								defer wg.Done()
								CreateMaster(patchfile, 2)
								mw.edit.SetText("Процедура завершена, вы можете продолжить работу.")
							}(patchfile, 3)
						} else {
							s := fmt.Sprintf("Сначала откройте подходящий файл отгрузки")
							mw.edit.SetText(s)
						}

					},
				}, PushButton{
					Text: "ТТН",
					OnClicked: func() {
						var wg sync.WaitGroup
						wg.Add(1)
						// закрываем в анонимной функции переменную из цикла,что бы предотвартить её потерю во время обработки
						mw.edit.SetText("Начался процесс создания TTN.")
						time.Sleep(15 * time.Millisecond)
						go func(mw *MyMainWindow) {
							defer wg.Done()
							TTN(mw)
							mw.edit.SetText("Процедура завершена, вы можете продолжить работу.")
						}(mw)

					},
				}},
			},
		},
	}

	if _, err := MW.Run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func TTN(mw *MyMainWindow) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println("405")
	}
	dlg := new(walk.FileDialog)

	dlg.FilePath = mw.path
	dlg.Title = "Select File"
	dlg.Filter = "Exe files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		mw.edit.AppendText("Error : File Open\r\n")
		return
	} else if !ok {
		mw.edit.AppendText("Cancel\r\n")
		return
	}
	//mw.path = dlg.FilePath
	ttnpatch := dlg.FilePath
	s := fmt.Sprintf("Select : %s\r\n", mw.path) + "Начался процесс создания ТТН\r\n."
	mw.edit.SetText(s)

	xlre, err := excelize.OpenFile(ttnpatch)
	if err != nil {
		fmt.Println("asdasdqwe")
		fmt.Println(err)
		return
	}
	d1 := xlre.GetCellValue("Sheet1", "A1")
	d2 := xlre.GetCellValue("Sheet1", "B1")
	d3 := xlre.GetCellValue("Sheet1", "C1")
	//	d4 := xlre.GetCellValue("Sheet1", "D1")
	os.MkdirAll(dir+"/TTN", os.ModePerm)
	data := [][]string{}
	for count := 3; len(xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count))) > 0; count++ {

		kmap := []string{"A", "C", "D", "E", "F", "G", "H", "J", "Y"}

		tdata := []string{}
		for _, value := range kmap {
			tdata = append(tdata, xlre.GetCellValue("Sheet1", value+strconv.Itoa(count)))

		}
		data = append(data, tdata)

		if len(xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count+1))) <= 0 {
			for cc := 1; cc < 10; cc++ {
				if len(xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count+cc))) > 0 {
					count = count + cc - 1
					break
				}

			}
		}
	}

	ccc := 0
	for _, x := range data {
		ccc++

		xlrt, err := excelize.OpenFile(dir + "/templ/ttnpl.xlsx")
		if err != nil {

			fmt.Println(err)
			return
		}
		sh := "Sheet1"
		xlrt.SetCellValue(sh, "A11", d2)
		xlrt.SetCellValue(sh, "A76", d3)
		xlrt.SetCellValue(sh, "A114", d3)
		xlrt.SetCellValue(sh, "BI8", x[0])
		xlrt.SetCellValue(sh, "BD43", x[0])
		xlrt.SetCellValue(sh, "A43", x[0])
		xlrt.SetCellValue(sh, "AD118", x[0])
		xlrt.SetCellValue(sh, "CF118", x[0])

		xlrt.SetCellValue(sh, "CM8", x[1])
		xlrt.SetCellValue(sh, "A16", x[5])
		xlrt.SetCellValue(sh, "A20", x[6]+" (м3)")
		xlrt.SetCellValue(sh, "BD47", x[6]+" (м3)")
		xlrt.SetCellValue(sh, "A47", x[6]+" (м3)")
		xlrt.SetCellValue(sh, "AR106", "Оплата доставки за "+x[7]+" (м3)")
		xlrt.SetCellValue(sh, "BD39", x[2])
		xlrt.SetCellValue(sh, "A39", x[8])
		xlrt.SetCellValue(sh, "BD11", d1)
		xlrt.SetCellValue(sh, "BN82", x[3])

		xlrt.SetCellValue(sh, "A78", x[4])
		xlrt.SetCellValue(sh, "BC118", "Водитель "+x[4])
		xlrt.SetCellValue(sh, "BD51", x[4])
		xlrt.SetCellValue(sh, "A51", x[4])

		if err != nil {
			fmt.Println("TTN/" + strconv.Itoa(ccc) + "-" + x[1] + ".xlsx")
			//return
		}

		err2 := xlrt.SaveAs(dir + "/TTN/" + strconv.Itoa(ccc) + "-" + x[1] + ".xlsx")
		if err2 != nil {
			fmt.Println(err2, dir+"/TTN/"+strconv.Itoa(ccc)+"-"+x[1]+".xlsx")
			return
		}

	}
}

func CreateMaster(patch string, flag int) {

	xlre, err := excelize.OpenFile(patch)
	if err != nil {
		fmt.Println(err)
		return
	}
	var dm map[string]map[string]map[string][][]string
	dm = make(map[string]map[string]map[string][][]string)
	OOO = make(map[string][]string)
	ccc := 2
	for len(xlre.GetCellValue("Info", "A"+strconv.Itoa(ccc))) > 0 {
		d1 := strings.Replace(strings.Replace(string(xlre.GetCellValue("Info", "A"+strconv.Itoa(ccc))), "/", "", -1), `"`, "", -1)
		d2 := xlre.GetCellValue("Info", "B"+strconv.Itoa(ccc))
		d3 := xlre.GetCellValue("Info", "C"+strconv.Itoa(ccc))

		OOO[d1] = []string{d2, d3}
		fmt.Println(OOO[d1])
		ccc++
	}
	for count := 3; len(xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count))) > 0; count++ {

		A := strings.Replace(strings.Replace(string(xlre.GetCellValue("Sheet1", "E"+strconv.Itoa(count))), "/", "", -1), `"`, "", -1)
		B := strings.Replace(strings.Replace(string(xlre.GetCellValue("Sheet1", "C"+strconv.Itoa(count))), "/", "", -1), `"`, "", -1)
		C := (xlre.GetCellValue("Sheet1", "A"+strconv.Itoa(count)))

		if len(B) <= 0 {
			B = "noname"
		}
		if len(A) <= 0 {
			A = "noname"
		}

		kmap := []string{"D", "G", "H", "I", "J", "K", "M", "N", "O", "P", "B", "S", "T", "R", "U", "V"}
		data := [][]string{}
		tdata := []string{}
		for _, value := range kmap {
			tdata = append(tdata, xlre.GetCellValue("Sheet1", value+strconv.Itoa(count)))

		}

		data = append(data, tdata)
		if len(dm[A]) == 0 {
			dm[A] = make(map[string]map[string][][]string)
		}
		if len(dm[A][B]) == 0 {
			dm[A][B] = make(map[string][][]string)
		}
		if len(dm[A][B][C]) == 0 {
			dm[A][B][C] = data
		} else {
			dm[A][B][C] = append(dm[A][B][C], tdata)
		}
	}

	for _, f2 := range reflect.ValueOf(dm).MapKeys() {
		firm := f2.Interface().(string)
		for _, url2 := range reflect.ValueOf(dm[firm]).MapKeys() {
			url := url2.Interface().(string)
			//fmt.Println(url)
			if flag == 1 {
				cxlsx(firm, url, dm[firm][url])
			}
			if flag == 2 {
				cclient(firm, url, dm[firm][url])
			}
			if flag == 3 {
				cxlsx(firm, url, dm[firm][url])
				cclient(firm, url, dm[firm][url])
			}

		}
	}
}

func pIntFloat(num string) float64 {
	if len(num) > 0 {
		rt, err := strconv.Atoi(num)
		if err != nil {
			rt, err2 := strconv.ParseFloat((num), 64)
			if err2 != nil {
				fmt.Println(err2, num)
				return 0
			} else {
				return rt
			}
		} else {
			return float64(rt)
		}
	} else {
		return 0
	}
	return 0
}

func cxlsx(from string, firm string, data map[string][][]string) {
	fmt.Println("Started work with Admins: " + firm)

	os.MkdirAll(fname+"/administrator/"+from, os.ModePerm)
	xlrt, err := excelize.OpenFile("templ/tmpl.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	sh := "Sheet1"
	if len(OOO[firm]) > 0 {
		xlrt.SetCellValue(sh, "A1", OOO[firm][0])

	}
	if len(OOO[from]) > 0 {
		xlrt.SetCellValue(sh, "B1", OOO[from][0])
		xlrt.SetCellValue(sh, "C1", OOO[from][1])
	}
	count := 3
	var keys []string
	for _, datef := range reflect.ValueOf(data).MapKeys() {

		datef3 := datef.Interface().(string)
		if len(datef3) > 0 {
			keys = append(keys, datef3[6:]+"-"+datef3[:2]+"-"+datef3[3:5])
		} else {
			keys = append(keys, "")
		}

	}
	sort.Strings(keys)
	for _, dateff := range keys {
		if len(dateff) > 5 {
			dateff = dateff[3:5] + "-" + dateff[6:] + "-" + dateff[:2]
		}
		for _, x := range data[dateff] {
			cc := strconv.Itoa(count)

			dateff2 := dateff[3:5] + "." + dateff[:2] + "." + dateff[6:]

			p1 := pIntFloat(x[6])
			p2 := pIntFloat(x[7])
			p3 := pIntFloat(x[8])
			p4 := pIntFloat(x[9])
			p5 := pIntFloat(x[10])
			if p1 != 0 {
				xlrt.SetCellValue(sh, "H"+cc, p1)
			}
			if p2 != 0 {
				xlrt.SetCellValue(sh, "J"+cc, p2)
			}
			if p3 != 0 {
				xlrt.SetCellValue(sh, "I"+cc, p3)
			}
			if p4 != 0 {
				xlrt.SetCellValue(sh, "K"+cc, p4)
			}
			if p5 != 0 {
				xlrt.SetCellValue(sh, "B"+cc, p5)
			}

			xlrt.SetCellValue(sh, "C"+cc, x[0])
			xlrt.SetCellValue(sh, "Y"+cc, x[1])
			xlrt.SetCellValue(sh, "E"+cc, x[3])
			xlrt.SetCellValue(sh, "F"+cc, x[4])
			xlrt.SetCellValue(sh, "G"+cc, x[5])

			xlrt.SetCellValue(sh, "A"+cc, dateff2)
			xlrt.SetCellValue(sh, "D"+cc, x[2])
			xlrt.SetCellValue(sh, "P"+cc, x[11])
			xlrt.SetCellValue(sh, "Q"+cc, x[12])
			xlrt.SetCellValue(sh, "R"+cc, x[13])
			xlrt.SetCellValue(sh, "U"+cc, x[14])
			xlrt.SetCellValue(sh, "W"+cc, x[15])

			count++
			if err != nil {
				fmt.Println(err, firm+".xlsx")
			}
		}
		count++

	}
	xlrt.UpdateLinkedValue()
	err2 := xlrt.SaveAs(fname + "/administrator/" + from + "/" + firm + ".xlsx")
	if err2 != nil {
		fmt.Println(err2, firm+".xlsx")
		return
	}
}

func cclient(from string, firm string, data map[string][][]string) {
	fmt.Println("Started work with client: " + firm)

	os.MkdirAll(fname+"/Clients/"+from, os.ModePerm)
	xlrt, err := excelize.OpenFile("templ/tmcl.xlsx")
	if err != nil {
		fmt.Println(err, "err1")
		return
	}

	sh := "Sheet1"
	if len(OOO[firm]) > 0 {
		xlrt.SetCellValue(sh, "A1", OOO[firm][0])
		xlrt.SetCellValue(sh, "B1", OOO[firm][1])
	}

	count := 3
	var keys []string
	for _, datef := range reflect.ValueOf(data).MapKeys() {
		datef3 := datef.Interface().(string)
		if len(datef3) > 0 {
			keys = append(keys, datef3[6:]+"-"+datef3[:2]+"-"+datef3[3:5])
		} else {
			keys = append(keys, "")
		}

	}
	sort.Strings(keys)
	for _, dateff := range keys {
		if len(dateff) > 5 {
			dateff = dateff[3:5] + "-" + dateff[6:] + "-" + dateff[:2]
		}
		for _, x := range data[dateff] {
			cc := strconv.Itoa(count)

			dateff2 := dateff[3:5] + "." + dateff[:2] + "." + dateff[6:]
			p1 := pIntFloat(x[6])
			p2 := pIntFloat(x[7])
			p3 := pIntFloat(x[8])
			p4 := pIntFloat(x[9])
			p5 := pIntFloat(x[10])
			if p1 != 0 {
				xlrt.SetCellValue(sh, "H"+cc, p1)
			}
			if p2 != 0 {
				xlrt.SetCellValue(sh, "J"+cc, p2)
			}
			if p3 != 0 {
				xlrt.SetCellValue(sh, "I"+cc, p3)
			}
			if p4 != 0 {
				xlrt.SetCellValue(sh, "K"+cc, p4)
			}
			if p5 != 0 {
				xlrt.SetCellValue(sh, "B"+cc, p5)
			}
			xlrt.SetCellValue(sh, "C"+cc, x[0])
			xlrt.SetCellValue(sh, "E"+cc, x[3])
			xlrt.SetCellValue(sh, "F"+cc, x[4])
			xlrt.SetCellValue(sh, "G"+cc, x[5])
			xlrt.SetCellValue(sh, "A"+cc, dateff2)
			xlrt.SetCellValue(sh, "D"+cc, x[2])

			count++

			if err != nil {
				fmt.Println(err, firm+".xlsx")
				//return
			}
		}
		count++

	}
	xlrt.UpdateLinkedValue()
	err2 := xlrt.SaveAs(fname + "/Clients/" + from + "/" + firm + ".xlsx")
	if err2 != nil {
		fmt.Println(err2, firm+".xlsx")
		return
	}

	//
}

func (mw *MyMainWindow) pbClicked() {
	dlg := new(walk.FileDialog)

	dlg.FilePath = mw.path
	dlg.Title = "Select File"
	dlg.Filter = "Exe files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		mw.edit.AppendText("Error : File Open\r\n")
		return
	} else if !ok {
		mw.edit.AppendText("Cancel\r\n")
		return
	}
	mw.path = dlg.FilePath
	patchfile = dlg.FilePath
	fname = mw.Name()
	temp := strings.Split(patchfile, "\\")
	temp2 := temp[len(temp)-1]
	fname = temp2[:len(temp2)-5]
	fmt.Println(fname)
	s := fmt.Sprintf("Select : %s\r\n", mw.path) + "Вы открыли этот файл отгрузки для работы. Выберите желаемое действие с этими данными.\r\n (Master\r\nMaster and Clients\r\nClients\r\n."
	mw.edit.SetText(s)
}

func _check(err error) {
	if err != nil {
		panic(err)
	}
}
