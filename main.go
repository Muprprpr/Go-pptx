// Package main provides an example usage of the go-pptx library
package main

import (
	"fmt"
	"os"

	"github.com/Muprprpr/Go-pptx/opc"
)

func main() {
	if len(os.Args) < 2 {
		printUsage()
		return
	}

	switch os.Args[1] {
	case "info":
		if len(os.Args) < 3 {
			fmt.Println("Usage: go-pptx info <file.pptx>")
			return
		}
		printInfo(os.Args[2])
	case "stream":
		if len(os.Args) < 3 {
			fmt.Println("Usage: go-pptx stream <file.pptx>")
			return
		}
		printInfoStream(os.Args[2])
	default:
		printUsage()
	}
}

func printUsage() {
	fmt.Println("Go-PPTX - PowerPoint file toolkit")
	fmt.Println()
	fmt.Println("Usage:")
	fmt.Println("  go-pptx info <file.pptx>     Show file info (traditional mode)")
	fmt.Println("  go-pptx stream <file.pptx>   Show file info (streaming mode)")
}

func printInfo(path string) {
	pkg, err := opc.OpenFile(path)
	if err != nil {
		fmt.Printf("Error opening file: %v\n", err)
		return
	}
	defer pkg.Close()

	fmt.Printf("File: %s\n", path)
	fmt.Printf("Parts: %d\n", pkg.PartCount())
	fmt.Println()

	// Print all parts
	for _, uri := range pkg.PartURIs() {
		part := pkg.GetPart(uri)
		fmt.Printf("  %s [%s]\n", uri.URI(), part.ContentType())
	}

	// Print relationships
	fmt.Println()
	fmt.Printf("Relationships: %d\n", pkg.Relationships().Count())
	for _, rel := range pkg.Relationships().All() {
		fmt.Printf("  %s -> %s\n", rel.RID(), rel.TargetURI().URI())
	}
}

func printInfoStream(path string) {
	pkg, err := opc.OpenStream(path)
	if err != nil {
		fmt.Printf("Error opening file: %v\n", err)
		return
	}
	defer pkg.Close()

	fmt.Printf("File: %s (streaming mode)\n", path)
	fmt.Printf("Parts: %d\n", pkg.PartCount())
	fmt.Println()

	// Print all parts (without loading content)
	for _, uri := range pkg.PartURIs() {
		part := pkg.GetPart(uri)
		fmt.Printf("  %s [%s] loaded=%v\n", uri.URI(), part.ContentType(), part.IsLoaded())
	}

	// Print relationships
	fmt.Println()
	fmt.Printf("Relationships: %d\n", pkg.Relationships().Count())
	for _, rel := range pkg.Relationships().All() {
		fmt.Printf("  %s -> %s\n", rel.RID(), rel.TargetURI().URI())
	}
}
