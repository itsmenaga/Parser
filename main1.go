package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/unidoc/unioffice/document"
)

func main() {
	// Check that the required command-line arguments were provided
	if len(os.Args) < 3 {
		fmt.Println("Usage: go run main.go input.docx output.csv")
		return
	}

	// Get the input and output filenames from the command-line arguments
	inputFilename := os.Args[1]
	outputFilename := os.Args[2]

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
		// Check if the paragraph is marked as "description", "impact", or "recommendation"
		style := para.Properties().ParagraphStyleId.String()
		if style == "description" || style == "impact" || style == "recommendation" {
			// Extract the text from the paragraph
			text := para.Text()

			// Write the text to the CSV file, including the paragraph style
			if err := w.Write([]string{style, text}); err != nil {
				fmt.Println("Error writing to CSV:", err)
				return
			}
		}
	}

	// Flush the CSV writer's buffer
	w.Flush()

	fmt.Println("Done!")
}
