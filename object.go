package e3

import (
	"errors"
	"fmt"

	ole "github.com/ZaytsveDmitriy/ole"
)

var errObject = errors.New("object error")

func toAnyTypeArr[T any](v *ole.VARIANT) (result []T) {
	array := v.ToArray().ToValueArray()

	if len(array) > 1 {
		for _, value := range array[1:] {
			if v, ok := value.(T); ok {
				result = append(result, v)
			} else {
				panic("Failed when type assertion")
			}
		}
	}

	return result
}

func getAnyArray[T any](o MethodCaller, methodName string) (result []T, cnt int32) {
	var data ole.VARIANT

	ole.VariantInit(&data)
	defer ole.VariantClear(&data)

	c, err := o.CallMethod(methodName, &data)
	if err != nil {
		fmt.Println("Error happened when call method:", methodName)
	}

	if cnt = c.Value().(int32); cnt > 0 {
		result = toAnyTypeArr[T](&data)
	}

	return result, cnt
}

type MethodCaller interface {
	CallMethod(string, ...interface{}) (*ole.VARIANT, error)
}

type object struct {
	*ole.IDispatch
}

func callMethod[T any](o MethodCaller, methodName string) (result T, err error) {
	res, err := o.CallMethod(methodName)
	defer res.Clear()

	if err != nil {
		log.Errorw("Failed to call method",
			"method name", methodName,
			"error description", err.Error())

		return result, fmt.Errorf("%w", err)
	}

	result, ok := res.Value().(T)
	if !ok {
		log.Errorw("Failed to call method",
			"method name", methodName,
			"error description", "Failed with return value type assertion")

		return result, fmt.Errorf("%w, %s", errObject, "Failed with return value type assertion")
	}

	return result, nil
}

func callMethodWithArgs[T any](o MethodCaller, methodName string, args ...interface{}) (result T, err error) {
	res, err := o.CallMethod(methodName, args...)
	defer res.Clear()

	if err != nil {
		log.Errorw("Failed to call method",
			"method name", methodName,
			"error description", err.Error())

		return result, fmt.Errorf("%w, Failed when call method %s", err, methodName)
	}

	result, ok := res.Value().(T)
	if !ok {
		log.Errorw("Failed to call method",
			"method name", methodName,
			"error description", "Failed with return value type assertion")

		return result, fmt.Errorf("%w, %s", errObject, "Failed with return value type assertion")
	}

	return result, nil
}

func (o *object) ID() (result int32) {
	const methodName = "GetId"

	result, err := callMethod[int32](o, methodName)
	if err != nil {
		return 0
	}

	return result
}

func (o *object) SetID(ID int32) (result int32) {
	const methodName = "SetId"

	result, err := callMethodWithArgs[int32](o, methodName, ID)
	if err != nil {
		return -1
	}

	return result
}

func (o *object) HyperlinkTextIDs() (IDs []int32, cnt int32) {
	IDs, cnt = getAnyArray[int32](o, "GetHyperlinkTextIds")

	return
}

func (o *object) Delete() {
	o.CallMethod("Delete")
}

func (o *object) DeleteAttribute(name string) {
	o.CallMethod("DeleteAttribte", name)
}

func (o *object) SetAttributeVisibility(name string) (result int32) {
	const methodName = "SetAttributeVisibility"

	result, err := callMethodWithArgs[int32](o, methodName, name, 1)
	if err != nil {
		result = -1
	}

	return result
}
