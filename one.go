package main

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/unidoc/unioffice/document"
)

func main() {
	// Open the Word document
	doc, err := document.Open("input.docx")
	if err != nil {
		fmt.Println("Error opening document:", err)
		return
	}

	// Create a new CSV file
	f, err := os.Create("output.csv")
	if err != nil {
		fmt.Println("Error creating CSV file:", err)
		return
	}
	defer f.Close()

	// Create a new CSV writer
	w := csv.NewWriter(f)

	// Iterate over the paragraphs in the document
	for _, para := range doc.Paragraphs() {
		// Check if the paragraph is marked as "description"
		if para.Properties().ParagraphStyleId.String() == "description" {
			// Extract the text from the paragraph
			text := para.Text()

			// Write the text to the CSV file
			if err := w.Write([]string{text}); err != nil {
				fmt.Println("Error writing to CSV:", err)
				return
			}
		}
	}

	// Flush the CSV writer's buffer
	w.Flush()

	fmt.Println("Done!")
}
