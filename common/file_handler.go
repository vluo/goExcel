package common

import (
	"crypto/rand"
	"math/big"
	"os"
	"strings"
)

func File_exists(filePath string) bool {
	_, err := os.Stat(filePath)
	if err == nil {
		return true
	}
	if os.IsNotExist(err) {
		return false
	}
	return true
}

func File_dir(filePath string) string {
	if filePath == "" {
		return ""
	}
	pathSeg := strings.Split(filePath, "/")
	if len(pathSeg) < 2 {
		return ""
	}
	return strings.Join(pathSeg[0:len(pathSeg)-1], "/")
}

func Rand_int(min, max int64) int64 {
	maxBigInt := big.NewInt(max)
	i, _ := rand.Int(rand.Reader, maxBigInt)
	if i.Int64() < min {
		Rand_int(min, max)
	}
	return i.Int64()
}
