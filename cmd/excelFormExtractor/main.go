package main

import (
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/championeng/excelFormExtractor/pkg/extractor"
)

func printSuccessAndExit(response extractor.Response) {
	jsonResponse, _ := json.Marshal(response)
	fmt.Println(string(jsonResponse))
	os.Exit(0)
}

func printErrorAndExit(response extractor.Response, code int) {
	jsonResponse, _ := json.Marshal(response)
	fmt.Fprintln(os.Stderr, string(jsonResponse))
	os.Exit(code)
}

func main() {
	// extractor.ReadFormControls()

	// Check if arguments are provided
	if len(os.Args) < 2 {
		// Return error for missing arguments
		response := extractor.Response{
			Status:  "error",
			Message: "No arguments provided",
		}
		printErrorAndExit(response, 1)
	}

	// Get the command from first argument
	command := os.Args[1]

	// Example command handling
	switch command {
	case "hello":
		// Success case
		response := extractor.Response{
			Status:  "success",
			Message: "Hello command executed successfully",
			Data:    "Hello, World!",
		}
		printSuccessAndExit(response)

	case "seccf":
		// Check for required second argument
		if len(os.Args) < 3 {
			response := extractor.Response{
				Status:  "error",
				Message: "SCEEF excel requires path name parameter",
			}
			printErrorAndExit(response, 2)
		}

		input := os.Args[2]

		seccf_extr, err := extractor.MakeSECCFExtractor(input, []string{"Amazon", "Amazon Inc", "Aamazon Ltd"})
		if err != nil {
			fmt.Printf("Failed to initialize extractor: %v\n", err)
			return
		}
		defer seccf_extr.Close()

		seccf_extr.Extract()
		// seccf_extr.ReadFormControls()

		// response := Response{
		// 	Status:  "success",
		// 	Message: "Processing completed",
		// 	Data:    map[string]string{"processed": strings.ToUpper(input)},
		// }
		// printSuccessAndExit(response)

	case "process":
		// Check for required second argument
		if len(os.Args) < 3 {
			response := extractor.Response{
				Status:  "error",
				Message: "Process command requires a parameter",
			}
			printErrorAndExit(response, 2)
		}

		// Process the input
		input := os.Args[2]
		if strings.ToLower(input) == "fail" {
			// Simulate a process failure
			response := extractor.Response{
				Status:  "error",
				Message: "Process failed as requested",
				Data:    map[string]string{"failed_input": input},
			}
			printErrorAndExit(response, 3)
		}

		// Success case with processed data
		response := extractor.Response{
			Status:  "success",
			Message: "Processing completed",
			Data:    map[string]string{"processed": strings.ToUpper(input)},
		}
		printSuccessAndExit(response)

	default:
		// Unknown command error
		response := extractor.Response{
			Status:  "error",
			Message: fmt.Sprintf("Unknown command: %s", command),
		}
		printErrorAndExit(response, 4)
	}
}
