# msscript
Microsoft JScript/VBScript Module for VFPx (includes CryptoJS as sample)

Usage: Run JavaScript/VBScript code in VFPx

```
** Initialize Module:
DO MSSCRIPT

** Encrypt String using CryptoJS:
EncryptedStr = MSSCR_EncryptAES("String to Encrypt","Encryption Key")

** Decrypt String using CryptoJS:
DecryptedStr = MSSCR_DecryptAES(EncryptedStr,"Encryption Key")

** Sample Function to run JavaScript Code:
FUNCTION JSFunc
PARAMETERS n1,n2

  ** STORE VARIABLE:
  PRIVATE o_Var = CREATEOBJECT("MSSCR_VAR")
  o_Var.AddProperty("n1",n1)
  o_Var.AddProperty("n2",n2)

  ** JAVASCRIPT CODE:
  PRIVATE c_Code
  TEXT TO c_Code NOSHOW
    _VAR.RetVal = _VAR.n1+_VAR.n2
  ENDTEXT

  ** RUN JAVASCRIPT CODE:
  PRIVATE RetVal
  RetVal = MSSCR_Exec(c_Code,o_Var)  && Run c_Code and update o_Var (as _VAR in JavaScript Code).
  RELEASE o_Var

RETURN (RetVal)
ENDFUNC

```
