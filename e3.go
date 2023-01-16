package e3

import (
	"fmt"

	ole "github.com/ZaytsveDmitriy/ole"
	"github.com/ZaytsveDmitriy/ole/oleutil"
)

type Logger interface {
	Infow(string, ...interface{})
	Errorw(string, ...interface{})
}

type E3 struct {
	object
}

// var log *zap.SugaredLogger
var log Logger

// func init() {
// 	prod, _ := zap.NewProduction()
// 	log = prod.Sugar()
// }

// New - create New instance of E3 object...
func New(logger Logger) *E3 {
	log = logger
	unknown, _ := oleutil.CreateObject("CT.Application")

	obj, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Infow("Fail to create E3")

		return nil
	}

	app := E3{object{obj}}

	return &app
}

func (e *E3) CreateNewJob() (j *Job) {
	obj, err := e.CallMethod("CreateJobObject")
	if err != nil {
		fmt.Println("No Job instance was returned")

		return nil
	}

	j = &Job{object{obj.ToIDispatch()}}

	return j
}

func (e *E3) Message(m string) {
	e.CallMethod("PutMessage", m)
}

func (e *E3) MessageEx(popUp bool, m string, item int32, red, green, blue int32) {
	e.CallMethod("PutMessageEx", popUp, m, item, red, green, blue)
}

func (e *E3) Info(popUp bool, m string) {
	e.CallMethod("PutInfo", popUp, m)
}

func (e *E3) Warning(popUp bool, m string) {
	e.CallMethod("PutWarning", popUp, m)
}

func (e *E3) Error(popUp bool, m string) {
	e.CallMethod("PutError", popUp, m)
}
