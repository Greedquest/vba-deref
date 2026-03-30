'assume presence of api declarations with signature property get/let memxxx

Public Property Get DerefIntrinsic(ByVal vt As VbVarType, ByVal ptr As LongPtr) As Variant
    memInt(VarPtr(DerefIntrinsic)) = vt
    memCur(VarPtr(DerefIntrinsic) + 8) = memCur(ptr)
End Property
