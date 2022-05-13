package main

import (
	"fmt"
	"github.com/Luxurioust/excelize"
	"github.com/shopspring/decimal"
	"log"
	"os"
	"os/signal"
	"strconv"
	"time"
)

type Teacher struct {
	name string
	lessonHour int
	salary int64
}

type Lesson struct {
	name string
	lessonHour int
}

type LessonSalary struct {
	name string
	lessonSalary int
}

type TeacherSalary struct {
	name string
	lessonHour int
	salary int64
	lessonSalary int

	salaryTmp int64
	lessonHourTmp int
}

func ReadExcelTeacher(filename string) map[string]Teacher {
	ret := map[string]Teacher{}
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println("读取excel文件出错", err.Error())
		return ret
	}

	sheets := f.GetSheetMap()
	fmt.Println(sheets)
	sheet1 := sheets[1]
	fmt.Println("第一个工作表", sheet1)
	rows, err := f.GetRows(sheet1)
	if err != nil {
		log.Println("读取excel sheet出错", err.Error())
		return ret
	}

	cols := []string{}

	for i, row := range rows {
		if i < 2 {
			continue
		}
		if i == 2 { //取得第一行的所有数据---execel表头
			for _, colCell := range row {
				cols = append(cols, colCell)

			}
			fmt.Println("列信息", cols)

		} else {
			teacher := Teacher{}
			for j, colCell := range row {
				if "姓名"== cols[j] {
					if colCell == "合计" {
						break
					}
					teacher.name = colCell
				} else if "课时数"== cols[j] {
					if 0 == len(colCell) {
						continue
					}

					if lessonHour, err := strconv.Atoi(colCell); err == nil {
						teacher.lessonHour = lessonHour
					} else {
						log.Println("【异常】课时数转换错误：", err.Error(), ", 课时数：", colCell)
					}

				} else if "课时工资"== cols[j] {
					if 0 == len(colCell) {
						continue
					}

					if salary, err := strconv.ParseFloat(colCell, 64); err == nil {
						teacher.salary = decimal.NewFromFloat(salary).Round(0).IntPart()
					} else {
						log.Println("【异常】课时工资转换错误：", err.Error(), ", 课时数：", colCell)
					}
				}
			}
			// fmt.Println("行：", i, ", 数据: ", teacher)
			if val, ok:=ret[teacher.name]; ok {
				fmt.Println("【异常】姓名重复：", val)
			} else {
				if len(teacher.name) > 0 {
					ret[teacher.name] = teacher
				}
			}
		}
	}

	return ret
}

func ReadExcelLesson(filename string) map[string]Lesson {
	ret := map[string]Lesson{}
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println("读取excel文件出错", err.Error())
		return ret
	}

	sheets := f.GetSheetMap()
	fmt.Println(sheets)
	sheet1 := sheets[1]
	fmt.Println("第一个工作表", sheet1)
	rows, err := f.GetRows(sheet1)
	if err != nil {
		log.Println("读取excel sheet出错", err.Error())
		return ret
	}

	cols := []string{}

	for i, row := range rows {
		if i < 1 {
			continue
		}
		if i == 1 { //取得第一行的所有数据---execel表头
			for _, colCell := range row {
				cols = append(cols, colCell)
			}
			fmt.Println("列信息", cols)

		} else {
			lesson := Lesson{}
			for j, colCell := range row {
				if "教师所属院系"== cols[j] {
					if "基础部" != colCell && "兼课" != colCell {
						break
					}
				} else if "姓名"== cols[j] {
					lesson.name = colCell
				} else if "课时总计"== cols[j] {
					if 0 == len(colCell) {
						continue
					}
					if intTmp, err := strconv.Atoi(colCell); err == nil {
						lesson.lessonHour = intTmp
					} else {
						log.Println("【异常】课时总计转换错误：", err.Error(), ", 课时数：", colCell)
					}
				}
			}

			if len(lesson.name) > 0 {
				if _, ok := ret[lesson.name]; !ok {
					ret[lesson.name] = lesson
				}
			}
		}
	}

	return ret
}

func ReadExcelLessonSalary(filename string) map[string]LessonSalary {
	ret := map[string]LessonSalary{}
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println("读取excel文件出错", err.Error())
		return ret
	}

	sheets := f.GetSheetMap()
	fmt.Println(sheets)
	sheet1 := sheets[1]
	fmt.Println("第一个工作表", sheet1)
	rows, err := f.GetRows(sheet1)
	if err != nil {
		log.Println("读取excel sheet出错", err.Error())
		return ret
	}

	cols := []string{}

	for i, row := range rows {
		if i == 0 { //取得第一行的所有数据---execel表头
			for _, colCell := range row {
				cols = append(cols, colCell)

			}
			fmt.Println("列信息", cols)

		} else {
			lessonSalary := LessonSalary{}
			for j, colCell := range row {
				if "姓名"== cols[j] {
					lessonSalary.name = colCell
				} else if "课时费"== cols[j] {
					// fmt.Println("课时费信息：", colCell)
					if 0 == len(colCell) {
						lessonSalary.lessonSalary = 0
					}  else {
						if intTmp, err := strconv.Atoi(colCell); err == nil {
							lessonSalary.lessonSalary = intTmp
						} else {
							log.Println("【异常】课时费转换错误：", err.Error(), ", 课时数：", colCell)
						}
					}
				}
			}
			// fmt.Println("行：", i, ", 数据: ", lessonSalary)
			if val, ok := ret[lessonSalary.name]; ok {
				log.Println("【异常】姓名重复：", val)
			} else {
				ret[lessonSalary.name] = lessonSalary
			}
		}
	}
	return ret
}

func GetMapKey(index int, indexChar string) string {
	return indexChar + strconv.Itoa(index)
}

func GetSalary(lessonHour int, lessonSalary int) int64 {
	var salary int

	if lessonHour <= 60 {
		salary = lessonHour * lessonSalary
	} else if lessonHour > 96 {
		salaryTmp := int(36.0 * float32(lessonSalary) * 1.2 + 0.5)
		salary = (lessonHour - 36) * lessonSalary + salaryTmp
	} else {
		salaryTmp := int(float32(lessonHour - 60) * float32(lessonSalary) * 1.2 + 0.5)
		salary = 60 * lessonSalary + salaryTmp
	}

	return int64(salary)
}

func WriteExcel(teacherSalary map[string]TeacherSalary) {
	fxlsx := excelize.NewFile()
	data := map[string]string{
		//列表名称
		"A1": "姓名",
		"B1": "课时费",
		"C1": "统计课时",
		"D1": "绩效课时",
		"E1": "计算工资",
		"F1": "绩效工资",
	}

	for k, v := range data {
		//设置单元格的值
		fxlsx.SetCellValue("Sheet1", k, v)
	}

	fmt.Println("********** 校验课时 **********")

	xlsxLesson := map[string]string{}
	index := 2

	for key := range teacherSalary {
		// 课时校验写表
		if  teacherSalary[key].lessonHour != teacherSalary[key].lessonHourTmp {
			fmt.Println("【课时异常】姓名", teacherSalary[key].name, ", 统计课时：", teacherSalary[key].lessonHourTmp, ", 工资表课时", teacherSalary[key].lessonHour)
		}
		if  teacherSalary[key].salary != teacherSalary[key].salaryTmp {
			fmt.Println("【工资异常】姓名", teacherSalary[key].name, ", 计算工资：", teacherSalary[key].salaryTmp, ", 原工资", teacherSalary[key].salary)
		}

		xlsxLesson[GetMapKey(index, "A")] = key
		xlsxLesson[GetMapKey(index, "B")] = strconv.Itoa(teacherSalary[key].lessonSalary)
		xlsxLesson[GetMapKey(index, "C")] = strconv.Itoa(teacherSalary[key].lessonHourTmp)
		xlsxLesson[GetMapKey(index, "D")] = strconv.Itoa(teacherSalary[key].lessonHour)
		xlsxLesson[GetMapKey(index, "E")] = strconv.FormatInt(teacherSalary[key].salaryTmp, 10)
		xlsxLesson[GetMapKey(index, "F")] = strconv.FormatInt(teacherSalary[key].salary, 10)
		index++
	}

	log.Println(xlsxLesson)

	for k, v := range xlsxLesson {
		//设置单元格的值
		fxlsx.SetCellValue("Sheet1", k, v)
	}

	/*for k, v := range xlsxSalary {
		//设置单元格的值
		fxlsx.SetCellValue("Sheet1", k, v)
		// style, _ := fxlsx.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#FFFF00"}}`)
		// style, _ := fxlsx.NewStyle(`{"background":{"color":"#777777"}}`)
		//style, _ := fxlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
		style, _ := fxlsx.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
		fxlsx.SetCellStyle("Sheet1", k, k, style)
	}*/

	err := fxlsx.SaveAs("./check_result.xlsx")
	if err != nil {
		log.Fatal(err)
	}
}

func main() {

	fmt.Println("********** 读取lesson_salary.xlsx **********")
	lessonSalaries := ReadExcelLessonSalary("excels/lesson_salary.xlsx")
	log.Print(lessonSalaries)

	fmt.Println("********** 读取lesson.xlsx **********")
	lessons := ReadExcelLesson("excels/lesson.xlsx")
	log.Println(lessons)

	fmt.Println("********** 读取teacher.xlsx **********")
	teachers := ReadExcelTeacher("excels/teacher.xlsx")
	log.Println(teachers)

	teacherSalary := map[string]TeacherSalary{}
	for key := range teachers {
		ts := TeacherSalary{}
		ts.name = key
		ts.salary = teachers[key].salary
		ts.lessonHour = teachers[key].lessonHour

		if val, ok := lessons[key]; ok {
			ts.lessonHourTmp = val.lessonHour
		} else {
			ts.lessonHourTmp = 0
			log.Println("姓名：",key, "，未录入课时")
		}

		if val, ok := lessonSalaries[key]; ok {
			ts.lessonSalary = val.lessonSalary
		} else {
			ts.lessonSalary = 0
			log.Println("姓名：",key, "，未录入课时费")
		}

		ts.salaryTmp = GetSalary(ts.lessonHourTmp, ts.lessonSalary)
		teacherSalary[key] = ts
	}

	log.Println("数据汇总完毕：",teacherSalary)
	WriteExcel(teacherSalary)

	c := make(chan os.Signal)
	signal.Notify(c)
	go func() {
		fmt.Println("Go routine running")
		time.Sleep(3*time.Second)
		fmt.Println("Go routine done")
	}()
	<-c
	fmt.Println("bye")
}