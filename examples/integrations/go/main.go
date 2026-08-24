package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
)

type response struct {
	OK    bool            `json:"ok"`
	Error json.RawMessage `json:"error"`
}

func main() {
	binary := os.Getenv("RUSDOX_BIN")
	if binary == "" {
		binary = "rusdox"
	}
	outputRoot := "target/integration-go"
	if len(os.Args) > 1 {
		outputRoot = os.Args[1]
	}
	outputRoot, err := filepath.Abs(outputRoot)
	if err != nil {
		panic(err)
	}
	request := map[string]any{
		"protocol_version": 1,
		"request_id":      "go-example",
		"operation":       "render",
		"source": map[string]any{
			"kind":    "inline",
			"format":  "yaml",
			"content": "version: 1\noutput_name: go-report\nblocks:\n  - type: title\n    text: Go integration\n  - type: body\n    text: The stable boundary is JSON, not a premature native SDK.\n",
		},
		"output": map[string]any{"directory": "go", "name": "go-report", "pdf": true},
	}
	payload, err := json.Marshal(request)
	if err != nil {
		panic(err)
	}
	payload = append(payload, '\n')
	command := exec.Command(binary, "serve", "stdio", "--output-root", outputRoot, "--max-requests", "1")
	command.Stdin = bytes.NewReader(payload)
	var stdout bytes.Buffer
	command.Stdout = &stdout
	command.Stderr = os.Stderr
	if err := command.Run(); err != nil {
		panic(err)
	}
	var result response
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		panic(err)
	}
	if !result.OK {
		panic(fmt.Sprintf("RusDox request failed: %s", result.Error))
	}
	fmt.Print(stdout.String())
}
