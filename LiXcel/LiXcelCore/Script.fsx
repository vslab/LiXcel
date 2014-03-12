// Per ulteriori informazioni su F#, visitare http://fsharp.net. Vedere il progetto 'Esercitazione su F#'
// per ulteriori linee guida sulla programmazione F#.

#r @"C:\Users\pq\Documents\GitHub\LiXcel\LiXcel\LiXcelCore\bin\Debug\LiXcelCore.dll"
open LiXcelCore

open System
open System.Threading
open System.Threading.Tasks

let formula = "=SUM!A3"

let q = Parser.tokenize formula
let q = Parser.parseExpr formula
// Viene definito qui il codice di script della libreria

let stoppalo = ref false

let load () =
    let rec lavoraSchiavo i =
        if !stoppalo then
            System.Console.WriteLine("stoppalo!")
        else if i < 0 then
            System.Console.WriteLine("lavoro finito")
        else
            System.Threading.Thread.Sleep (1000)
            System.Console.WriteLine("task {0}", i)
            lavoraSchiavo (i-1)
    lavoraSchiavo 10

let t = new System.Threading.Thread(load)
t.Start()
stoppalo := true

      

