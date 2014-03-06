namespace  LiXcelCore
open Microsoft.Office.Interop

type FunctionLibrary =
    static member rng = new System.Security.Cryptography.RNGCryptoServiceProvider()//new System.Random()
    static member SUM values =
        List.sum<float> values
    static member LXLBYTES count =
        let rn = Array.zeroCreate count
        FunctionLibrary.rng.GetBytes(rn)
        rn
    static member LXLUNIFORM left right =
        let rn = FunctionLibrary.LXLBYTES 8
        let a = uint64(System.BitConverter.ToInt64(rn,0))
        let a = a ||| 0x8000000000000000UL// 1<<63
        (((float a) / (float 0x8000000000000000UL))-1.0) * (right-left) + left;
    static member LXLGAUSS m sigma = 
        let u = FunctionLibrary.LXLUNIFORM 1.0 0.0
        let v = FunctionLibrary.LXLUNIFORM 0.0 1.0
        m + sigma * sqrt(-2. * log(u)) * cos(2. * System.Math.PI * v)
    static member LN x = log x
    static member PI = System.Math.PI
    static member COS x = cos x
    static member SQRT x = sqrt x
    static member ROUND number digits = let mul = (10.0**truncate(digits)) in (truncate (number * mul + 0.5)) / mul