namespace  LiXcelCore
open Microsoft.Office.Interop

type FunctionLibrary =
    static member rng = new System.Random()
    static member SUM values =
        List.sum<float> values
    static member LXLUNIFORM left right =
        (FunctionLibrary.rng.NextDouble()) * (right - left) + left
    static member LXLGAUSS m sigma = 
        let u = FunctionLibrary.LXLUNIFORM 1.0 0.0
        let v = FunctionLibrary.LXLUNIFORM 0.0 1.0
        m + sigma * sqrt(-2. * log(u)) * cos(2. * System.Math.PI * v)