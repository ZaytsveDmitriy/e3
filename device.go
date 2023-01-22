package e3

// native method names
const (
	IsAssembly      = "IsAssembly"
	IsTerminalBlock = "IsTerminalBlock"
	IsTerminal      = "IsTerminal"
	IsAssemblyPart  = "IsAssemblyPart"
	IsBlock         = "IsBlock"
	IsDevice        = "IsDevice"
	IsConnector     = "IsConnector"
	IsCable         = "IsCable"
	IsWiregroup     = "IsWiregroup"
	IsMount         = "IsMount"
	IsCableDuct     = "IsCableDuct"
	IsHose          = "IsHose"
	IsTube          = "IsTube"
)

type Device struct {
	object
}

func (d *Device) RefDes() string {
	const methodName = "GetName"

	refDes, err := callMethod[string](d, methodName)
	if err != nil {
		return ""
	}

	return refDes
}

func boolMethod(o MethodCaller, methodName string) bool {
	res, err := callMethod[int32](o, methodName)
	if err != nil {
		return false
	}

	return res == 1
}

func (d *Device) IsDevice() bool {
	return boolMethod(d, IsDevice)
}

func (d *Device) IsAssembly() bool {
	return boolMethod(d, IsAssembly)
}

func (d *Device) IsAssemblyPart() bool {
	return boolMethod(d, IsAssemblyPart)
}

func (d *Device) IsBlock() bool {
	return boolMethod(d, IsBlock)
}

func (d *Device) IsCable() bool {
	return boolMethod(d, IsCable)
}

func (d *Device) IsCableDuct() bool {
	return boolMethod(d, IsCableDuct)
}

func (d *Device) IsConnector() bool {
	return boolMethod(d, IsConnector)
}

func (d *Device) IsHose() bool {
	return boolMethod(d, IsHose)
}

func (d *Device) IsMount() bool {
	return boolMethod(d, IsMount)
}

func (d *Device) IsTerminal() bool {
	return boolMethod(d, IsTerminal)
}

func (d *Device) IsTerminalBlock() bool {
	return boolMethod(d, IsTerminalBlock)
}

func (d *Device) IsWiregroup() bool {
	return boolMethod(d, IsWiregroup)
}

func (d *Device) IsTube() bool {
	return boolMethod(d, IsTube)
}
