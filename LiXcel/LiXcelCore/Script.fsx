// Per ulteriori informazioni su F#, visitare http://fsharp.net. Vedere il progetto 'Esercitazione su F#'
// per ulteriori linee guida sulla programmazione F#.

#I @"C:\Users\pq\Documents\GitHub\LiXcel\LiXcel\LiXcelCore\bin\Debug"
#r @"LiXcelCore.dll"
open LiXcelCore

let formula = "=SUM!A3"

let q = Parser.tokenize formula
let q = Parser.parseExpr formula
// Viene definito qui il codice di script della libreria

