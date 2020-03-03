// +build windows

package ole

import (
	"fmt"
	"unsafe"

	"github.com/pkg/errors"
)

func safeArrayFromByteSlice(slice []byte) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_UI1, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []byte to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []byte")
		}
	}
	return array, nil
}

func safeArrayFromStringSlice(slice []string) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []string to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(SysAllocStringLen(v)))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []string")
		}
	}
	return array, nil
}

func safeArrayFromUint16Slice(slice []uint16) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_UI2, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []uint16 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []uint16")
		}
	}
	return array, nil
}

func safeArrayFromUint32Slice(slice []uint32) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_UI4, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []uint32 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []uint32")
		}
	}
	return array, nil
}

func safeArrayFromUint64Slice(slice []uint64) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_UI8, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []uint64 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []uint64")
		}
	}
	return array, nil
}

func safeArrayFromInt16Slice(slice []int16) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_I2, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []int16 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []int16")
		}
	}
	return array, nil
}

func safeArrayFromInt32Slice(slice []int32) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_I4, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []int32 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []int32")
		}
	}
	return array, nil
}

func safeArrayFromInt64Slice(slice []int64) (*SafeArray, error) {
	array, _ := safeArrayCreateVector(VT_I8, 0, uint32(len(slice)))

	if array == nil {
		return nil, fmt.Errorf("Could not convert []int64 to SAFEARRAY")
	}

	for i, v := range slice {
		if err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v))); err != nil {
			return nil, errors.Wrap(err, "safeArrayPutElement for []int64")
		}
	}
	return array, nil
}
