GOOS=$(shell go env GOOS)

build:
	CGO_ENABLED=0 GOOS=$(GOOS) go build -ldflags="-s -w" -o ./xlsx-splitter.bin cmd/main.go

minify:
	upx -9 ./xlsx-splitter.bin
