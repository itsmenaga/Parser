package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/unidoc/unioffice/document"
)

func main() {
	// Check that the required command-line arguments were provided
	if len(os.Args) < 4 {
		fmt.Println("Usage: go run main.go input.docx output.csv style1 style2 ...")
		return
	}

	// Get the input and output filenames from the command-line arguments
	inputFilename := os.Args[1]
	outputFilename := os.Args[2]

	// Get the list of styles to parse from the command-line arguments
	styles := os.Args[3:]

	// Open the Word document
	doc, err := document.Open(inputFilename)
	if err != nil {
		fmt.Println("Error opening document:", err)
		return
	}

	// Create a new CSV file
	f, err := os.Create(outputFilename)
	if err != nil {
		fmt.Println("Error creating CSV file:", err)
		return
	}
	defer f.Close()

	// Create a new CSV writer
	w := csv.NewWriter(f)

	// Iterate over the paragraphs in the document
	for _, para := range doc.Paragraphs() {
		// Check if the paragraph is marked with one of the specified styles
		style := para.Properties().ParagraphStyleId.String()
		for _, s := range styles {
			if style == s {
				// Extract the text from the paragraph
				text := para.Text()

				// Write the text to the CSV file, including the paragraph style
				if err := w.Write([]string{style, text}); err != nil {
					fmt.Println("Error writing to CSV:", err)
					return
				}
				break
			}
		}
	}

	// Flush the CSV writer's buffer
	w.Flush()

	fmt.Println("Done!")
}
