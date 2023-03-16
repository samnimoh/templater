package main

import (
	"fmt"
	//"os"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/widget"
	docx "github.com/lukasjarosch/go-docx"
	"github.com/xuri/excelize/v2"
)

func main() {
	// create a new Fyne application
	myApp := app.New()

	// create a new window with a fixed size
	myWindow := myApp.NewWindow("Templater: by Samuel Nimoh and ChatGPT")
	myWindow.Resize(fyne.NewSize(800, 800))

	// create the label and button for the "Template" file
	templateLabel := widget.NewLabel("Template:")
	templateName := widget.NewLabel("")
	templateButton := widget.NewButton("Upload", func() {
		fileDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err == nil && reader != nil {
				// set the file name label to the uploaded file's name
				templateName.SetText(reader.URI().Path())

				// do something with the uploaded file
				fmt.Printf("Uploaded file: %s\n", reader.URI().String())
				reader.Close()
			}
		}, myWindow)
		//fileDialog.SetFilter([]string{"*.docx"})
		fileDialog.Show()
	})

	// create the label and button for the "Variables" file
	variablesLabel := widget.NewLabel("Variables:")
	variablesName := widget.NewLabel("")
	variablesButton := widget.NewButton("Upload", func() {
		fileDialog := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err == nil && reader != nil {
				// set the file name label to the uploaded file's name
				variablesName.SetText(reader.URI().Path())

				// do something with the uploaded file
				fmt.Printf("Uploaded file: %s\n", reader.URI().Path())
				reader.Close()
			}
		}, myWindow)
		//fileDialog.SetFilter([]string{"*.xlsx"})
		fileDialog.Show()
	})

	// create the button to generate the files
	generateButton := widget.NewButton("Generate Files", func() {
		variablesFile, err := excelize.OpenFile(variablesName.Text)
		if err != nil {
			fmt.Println(err)
			return
		}
		rows, err := variablesFile.GetRows("Sheet1")
		keys := rows[0]
		fmt.Println(rows)
		values := rows[1:]
		fmt.Println(values)

		templateDoc, err := docx.Open(templateName.Text)
		if err != nil {
			fmt.Println(err)
			return
		}

		for _, row := range values {
			replaceMap := make(docx.PlaceholderMap)
			for i, key := range keys {
				replaceMap[key] = row[i]
			}

			err := templateDoc.ReplaceAll(replaceMap)
			if err != nil {
				fmt.Println(replaceMap)
				fmt.Println("before error 1")
				fmt.Println(err)
				return
			}

			newDocName := fmt.Sprintf("%s.docx", row[0])
			err = templateDoc.WriteToFile(newDocName)
			if err != nil {
				fmt.Println("before error 2")
				fmt.Println(err)
				return
			}
			fmt.Printf("Generated file: %s\n", newDocName)
		}
	})

	// create the layout for the window
	content := container.NewVBox(
		container.NewHBox(templateLabel, templateName, templateButton),
		container.NewHBox(variablesLabel, variablesName, variablesButton),
		generateButton,
	)

	// set the window content and show it
	myWindow.SetContent(content)
	myWindow.ShowAndRun()
}
