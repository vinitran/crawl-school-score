package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"net/http"
	"sync"
)

type Score struct {
	SchoolCode   string `json:"schoolCode"`
	SchoolName   string `json:"schoolName"`
	Score        string `json:"score"`
	MajorsCode   string `json:"majorsCode"`
	MajorsName   string `json:"majorsName"`
	ProvinceName string `json:"provinceName"`
	SubjectGroup string `json:"subjectGroup"`
	SchoolSlug   string `json:"schoolSlug"`
	MajorsSlug   string `json:"majorsSlug"`
}

type ScoreResponse struct {
	Status     bool   `json:"status"`
	ErrorCode  int    `json:"errorCode"`
	Messages   []any  `json:"messages"`
	Version    int    `json:"version"`
	ServerTime string `json:"serverTime"`
	Data       struct {
		Model struct {
			Provinces []struct {
				Code        string `json:"code"`
				Name        string `json:"name"`
				Priority    int    `json:"priority"`
				PagingItems any    `json:"pagingItems"`
			} `json:"provinces"`
			SchoolTypes   []int   `json:"schoolTypes"`
			Scores        []Score `json:"scores"`
			Years         []int   `json:"years"`
			SubjectGroups []struct {
				Name string `json:"name"`
				Code string `json:"code"`
			} `json:"subjectGroups"`
			MaxScore    int    `json:"maxScore"`
			Year        int    `json:"year"`
			Province    any    `json:"province"`
			From        string `json:"from"`
			To          string `json:"to"`
			Group       string `json:"group"`
			Type        int    `json:"type"`
			PageIndex   int    `json:"pageIndex"`
			PageSize    int    `json:"pageSize"`
			TotalRow    int    `json:"totalRow"`
			PagingItems any    `json:"pagingItems"`
		} `json:"model"`
	} `json:"data"`
	ServerTimeUtc string `json:"serverTimeUtc"`
	DataEncrypt   any    `json:"dataEncrypt"`
	KeyEncrypt    any    `json:"keyEncrypt"`
}

const (
	API_URL       = `https://vietnamnet.vn/newsapi/EducationScore/GetSchoolByScore?componentId=COMPONENT002310&from=0&group=A&pageId=499756218a1449d9b9305de4c14db9bb&pageIndex=%d&pageSize=20&to=40&type=2&year=%d`
	maxGoroutines = 1
)

func main() {
	fileWrite := excelize.NewFile()

	err := fileWrite.SetCellStr("Sheet1", fmt.Sprintf("A%d", 1), "School Code")
	if err != nil {
		log.Fatal(err)
	}

	err = fileWrite.SetCellStr("Sheet1", fmt.Sprintf("B%d", 1), "School Name")
	if err != nil {
		log.Fatal(err)
	}

	err = fileWrite.SetCellStr("Sheet1", fmt.Sprintf("C%d", 1), "Major Code")
	if err != nil {
		log.Fatal(err)
	}

	err = fileWrite.SetCellStr("Sheet1", fmt.Sprintf("D%d", 1), "Major Name")
	if err != nil {
		log.Fatal(err)
	}

	err = fileWrite.SetCellStr("Sheet1", fmt.Sprintf("E%d", 1), "Subject Group")
	if err != nil {
		log.Fatal(err)
	}

	err = fileWrite.SetCellStr("Sheet1", fmt.Sprintf("F%d", 1), "Score")
	if err != nil {
		log.Fatal(err)
	}

	guard := make(chan struct{}, maxGoroutines)
	waitGroup := sync.WaitGroup{}

	rowExel := 2

	for page := 0; page < 200; page++ {
		log.Println("page:", page)
		guard <- struct{}{}
		waitGroup.Add(1)
		go func(file *excelize.File, page, row int) {
			api := fmt.Sprintf(API_URL, page, 2015)
			response, err := http.Get(api)
			if err != nil {
				log.Println(err)
				<-guard
				waitGroup.Done()
				return
			}

			responseData, err := io.ReadAll(response.Body)
			if err != nil {
				log.Println(err)
				<-guard
				waitGroup.Done()
				return
			}

			var responseObject ScoreResponse
			json.Unmarshal(responseData, &responseObject)
			if len(responseObject.Data.Model.Scores) == 0 {
				<-guard
				waitGroup.Done()
				return
			}

			if responseObject.Data.Model.Scores == nil {
				<-guard
				waitGroup.Done()
				return
			}
			err = ExportDataToExel(file, row, responseObject.Data.Model.Scores)
			if err != nil {
				log.Println(err)
				<-guard
				waitGroup.Done()
				return
			}
			<-guard
			waitGroup.Done()
		}(fileWrite, page, rowExel)

		rowExel = rowExel + 20
	}
	waitGroup.Wait()
}

func ExportDataToExel(fileWrite *excelize.File, row int, scores []Score) error {
	for i, data := range scores {
		score := []interface{}{data.SchoolCode, data.SchoolName, data.MajorsCode, data.MajorsName, data.SubjectGroup, data.Score}
		err := fileWrite.SetSheetRow("Sheet1", fmt.Sprintf("A%d", row+i), &score)
		if err != nil {
			return err
		}
	}

	err := fileWrite.SaveAs("dataScore2015.xlsx")
	if err != nil {
		return err
	}
	return nil
}
