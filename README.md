# Go OLE + possible []int32 input
이 패키지는 [github.com/go-ole/go-ole](https://github.com/go-ole/go-ole)를 가져와 입력값으로 []int32를 사용할 수 있도록 idispatch_windows.go와 safearrayslices.go를 수정한 것입니다.


## 사용법
설치
```
go get github.com/hspan/go-ole
```

상세한 사용법은 github.com/hspan/go-ole 와 동일하므로 [godoc](https://godoc.org/github.com/go-ole/go-ole)을 참조하시기 바랍니다.

## 수정내용
### idispatch_windows.go - invoke함수 내 case 추가
```go
func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
    .
    .
    .
			case []string:
				safeByteArray := safeArrayFromStringSlice(v.([]string))
				vargs[n] = NewVariant(VT_ARRAY|VT_BSTR, int64(uintptr(unsafe.Pointer(safeByteArray))))
                defer VariantClear(&vargs[n])

            // 추가 부분 시작 
			case []int32:  
				safeByteArray := safeArrayFromInt32Slice(v.([]int32))
                vargs[n] = NewVariant(VT_ARRAY|VT_I4, int64(uintptr(unsafe.Pointer(safeByteArray))))
            // 추가 부분 종료

				defer VariantClear(&vargs[n])
			default:
				panic("unknown type")
			}
		}
.
.
.
	return
}
```

### safearrayslices.go - 아래 함수 추가
```go
func safeArrayFromInt32Slice(slice []int32) *SafeArray {
	array, _ := safeArrayCreateVector(VT_I4, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []int32 to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}
```