'assume presence of api declarations with signature property get/let memxxx
'assume Declare Function VariantCopy Lib "oleaut32" (ByVal pDst As LongPtr, ByVal pSrc As LongPtr) As Long

Public Property Get DerefIntrinsic(ByVal vt As VbVarType, ByVal ptr As LongPtr) As Variant
    Dim src As Variant
    memInt(VarPtr(src)) = vt
    memCur(VarPtr(src) + 8) = memCur(ptr)
    VariantCopy VarPtr(DerefIntrinsic), VarPtr(src)
    memInt(VarPtr(src)) = vbEmpty 'prevent VariantClear double-free on scope exit
End Property
