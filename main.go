package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

func main() {
	// Open the source Excel file
	sourceFilePath := "cash.xlsx"
	srcFile, err := xlsx.OpenFile(sourceFilePath)
	if err != nil {
		log.Fatal(err)
	}

	// Create a new Excel file
	destinationFilePath := xlsx.NewFile()

	// Specify the sheet names to copy
	sheetNames := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}

	// Iterate over the source sheets and copy them to the destination file
	for _, sheetName := range sheetNames {
		srcSheet := srcFile.Sheet[sheetName]
		if srcSheet == nil {
			log.Printf("Sheet '%s' not found in the source file", sheetName)
			continue
		}

		// Create a new sheet in the destination file
		destinationSheet, err := destinationFilePath.AddSheet(sheetName)
		if err != nil {
			log.Fatalf("Failed to add sheet '%s' to the destination file: %v", sheetName, err)
		}

		// Iterate over the source rows and copy them to the destination sheet
		for _, srcRow := range srcSheet.Rows {
			dstRow := destinationSheet.AddRow()

			// Iterate over the source cells and copy them to the destination row
			for _, srcCell := range srcRow.Cells {
				dstCell := dstRow.AddCell()

				// Copy the cell value and formatting
				dstCell.Value = srcCell.Value
				dstCell.SetStyle(srcCell.GetStyle())
			}
		}
	}

	// Save the destination Excel file
	err = destinationFilePath.Save("destination.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Sheets copied successfully!")
}

// **********************************************************************************************//

// package main

// import (
// 	"fmt"
// 	"log"

// 	"github.com/tealeg/xlsx"
// )

// func main() {
// 	// Open the source Excel file
// 	sourceFilePath := "cash.xlsx"
// 	srcFile, err := xlsx.OpenFile(sourceFilePath)
// 	if err != nil {
// 		log.Fatal(err)
// 	}

// 	// Create a new Excel file
// 	destinationFilePath := xlsx.NewFile()

// 	// Specify the sheet names to copy
// 	sheetNames := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}

// 	// Iterate over the source sheets and copy them to the destination file
// 	for _, sheetName := range sheetNames {
// 		srcSheet := srcFile.Sheet[sheetName]
// 		if srcSheet == nil {
// 			log.Printf("Sheet '%s' not found in the source file", sheetName)
// 			continue
// 		}

// 		// Create a new sheet in the destination file
// 		destinationSheet, err := destinationFilePath.AddSheet(sheetName)
// 		if err != nil {
// 			log.Fatalf("Failed to add sheet '%s' to the destination file: %v", sheetName, err)
// 		}

// 		// Iterate over the source rows and copy them to the destination sheet
// 		for _, srcRow := range srcSheet.Rows {
// 			dstRow := destinationSheet.AddRow()

// 			// Iterate over the source cells and copy them to the destination row
// 			for _, srcCell := range srcRow.Cells {
// 				dstCell := dstRow.AddCell()

// 				// Copy the cell value
// 				dstCell.Value = srcCell.Value

// 				// Copy the cell style
// 				dstCell.SetStyle(copyCellStyle(srcCell.GetStyle(), destinationFilePath))

// 			}
// 		}
// 	}

// 	// Save the destination Excel file
// 	err = destinationFilePath.Save("destination.xlsx")
// 	if err != nil {
// 		log.Fatal(err)
// 	}

// 	fmt.Println("Sheets copied successfully!")
// }

// // Helper function to copy cell style from source to destination file
// func copyCellStyle(srcStyle *xlsx.Style, dstFile *xlsx.File) *xlsx.Style {
// 	if srcStyle == nil {
// 		return nil
// 	}

// 	dstStyle := xlsx.NewStyle()

// 	dstStyle = srcStyle

// 	return dstStyle
// }

// **********************************************************************************************//

// package main

// import (
// 	"fmt"
// 	"log"

// 	"github.com/tealeg/xlsx"
// )

// func main() {
// 	// Define the names of the Excel sheets you want to copy
// 	sheetsToCopy := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}

// 	// Define the path of the source zipped Excel file
// 	sourceFilePath := "cash.xlsx"

// 	// Define the path of the destination zipped Excel file
// 	destinationFilePath := "destination.xlsx"

// 	// Open the source zipped Excel file
// 	sourceFile, err := xlsx.OpenFile(sourceFilePath)
// 	if err != nil {
// 		log.Fatalf("Failed to open source file: %v", err)
// 	}

// 	// Create a new Excel file for the destination
// 	destinationFile := xlsx.NewFile()

// 	// Copy data from each source sheet to the destination
// 	for _, sheetName := range sheetsToCopy {
// 		// Get the source sheet by name
// 		sourceSheet, found := sourceFile.Sheet[sheetName]
// 		if !found {
// 			log.Printf("Sheet '%s' not found in the source file", sheetName)
// 			continue
// 		}

// 		// Create a new sheet in the destination file
// 		destinationSheet, err := destinationFile.AddSheet(sheetName)
// 		if err != nil {
// 			log.Printf("Failed to create sheet '%s' in the destination file: %v", sheetName, err)
// 			continue
// 		}

// 		// Copy data from the source sheet to the destination sheet
// 		err = copySheetData(sourceSheet, destinationSheet)
// 		if err != nil {
// 			log.Printf("Failed to copy data from sheet '%s': %v", sheetName, err)
// 		}
// 	}

// 	// Save the destination file
// 	err = destinationFile.Save(destinationFilePath)
// 	if err != nil {
// 		log.Fatalf("Failed to save destination file: %v", err)
// 	}

// 	fmt.Println("Data copied successfully!")
// }

// // Helper function to copy data from a source sheet to a destination sheet
// func copySheetData(sourceSheet, destinationSheet *xlsx.Sheet) error {
// 	// Copy cell data
// 	for _, sourceRow := range sourceSheet.Rows {
// 		destinationRow := destinationSheet.AddRow()
// 		for _, sourceCell := range sourceRow.Cells {
// 			destinationCell := destinationRow.AddCell()

// 			// Copy the cell value
// 			destinationCell.Value = sourceCell.Value

// 			// Copy the cell style
// 			destinationCell.SetStyle(sourceCell.GetStyle())
// 		}
// 	}

// 	return nil
// }
