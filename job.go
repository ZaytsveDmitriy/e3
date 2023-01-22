package e3

import (
	"fmt"

	ole "github.com/ZaytsveDmitriy/ole"
)

// Job ...
type Job struct {
	object
}

// NewSheet ...
func (j *Job) NewSheet() *Sheet {
	obj, err := j.CallMethod("CreateSheetObject")
	if err != nil {
		log.Errorw("Can't create new sheet")
	}

	sheet := Sheet{object{obj.ToIDispatch()}}

	return &sheet
}

func (j *Job) NewDevice() *Device {
	obj, err := j.CallMethod("CreateDeviceObject")
	if err != nil {
		log.Errorw("Can't create new device")
	}

	dev := Device{object{obj.ToIDispatch()}}

	return &dev
}

// AvailableLanguages ...
func (j *Job) AvailableLanguages() (languages []string, cnt int32) {
	languages, cnt = getAnyArray[string](j, "GetAvailableLanguages")

	return languages, cnt
}

// Languages ...
func (j *Job) Languages() (languages []string) {
	var data ole.VARIANT

	ole.VariantInit(&data)
	defer ole.VariantClear(&data)

	_, err := j.CallMethod("GetLanguages", &data)
	if err != nil {
		fmt.Println("Error happened when call method: GetLanguages", "error is", err)
	}

	if data.VT != ole.VT_EMPTY {
		languages = toAnyTypeArr[string](&data)
	}

	return
}

// SetLanguages ...
func (j *Job) SetLanguages(languages []string) {
	const methodName = "SetLanguages"

	_, err := j.CallMethod(methodName, languages)
	if err != nil {
		log.Errorw(
			"Failed to call method",
			"Method name", methodName,
			"Error description", fmt.Sprintf("language set : %s", languages),
		)
	}
}

// Levels ...
func (j *Job) Levels() (levels []int32) {
	var data, data1 ole.VARIANT

	ole.VariantInit(&data)
	defer ole.VariantClear(&data)

	ole.VariantInit(&data1)
	defer ole.VariantClear(&data1)

	_, err := j.CallMethod("GetLevels", &data, &data1, &data1, &data1, &data1, &data1, &data1)
	if err != nil {
		fmt.Println("Error happened when call method: GetLevels")
	}

	levels = toAnyTypeArr[int32](&data)

	return levels
}

// SetLevel ...
func (j *Job) SetLevel(n int32) error {
	const methodName = "SetLevel"

	_, err := callMethodWithArgs[int32](j, methodName, n, 1)
	if err != nil {
		return err
	}

	return nil
}

// ResetLevel ...
func (j *Job) ResetLevel(n int32) error {
	const methodName = "SetLevel"

	_, err := callMethodWithArgs[int32](j, methodName, n, 0)
	if err != nil {
		return err
	}

	return nil
}

// ExportPDF - make PDF file. options is a bitmask ...
func (j *Job) ExportPDF(fileName string, sheetIDs []int32, options uint16) (result int32) {
	resultVAR, err := j.CallMethod("ExportPDF", fileName, sheetIDs, options)
	if err != nil {
		log.Errorw("Failed when call Export PDF",
			"filename", fileName,
			"sheetIDs", sheetIDs,
			"options", options,
		)

		return -1
	}

	result, ok := resultVAR.Value().(int32)
	if !ok {
		panic("Faild when type assertion")
	}

	log.Infow("PDF was created successfully", "filename", fileName)

	return result
}

// SaveAs ...
func (j *Job) SaveAs(fileName string) (result int32) {
	const methodName = "SaveAs"

	result, err := callMethodWithArgs[int32](&j.object, methodName, fileName)
	if err != nil {
		result = -1
	}

	return result
}

// SelectedSymbolIds ...
func (j *Job) SelectedSymbolIds() (IDs []int32, cnt int32) {
	IDs, cnt = getAnyArray[int32](j, "GetSelectedSymbolIds")

	return IDs, cnt
}

// TreeSelectedSheetIds ...
func (j *Job) TreeSelectedSheetIds() (IDs []int32, cnt int32) {
	IDs, cnt = getAnyArray[int32](&j.object, "GetTreeSelectedSheetIds")

	return IDs, cnt
}

// TreeSelectedSheetIDsByFolder ...
func (j *Job) TreeSelectedSheetIDsByFolder() (IDs []int32, cnt int32) {
	IDs, cnt = getAnyArray[int32](&j.object, "GetTreeSelectedSheetIdsByFolder")

	return IDs, cnt
}

func (j *Job) TreeSelectedDeviceIds() (IDs []int32, cnt int32) {
	IDs, cnt = getAnyArray[int32](j, "GetTreeSelectedAllDeviceIds")

	return IDs, cnt
}
