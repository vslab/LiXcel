namespace  LiXcelCore
open Microsoft.Office.Interop

type FunctionLibrary =
    static member SUM values =
        List.sum<float> values
