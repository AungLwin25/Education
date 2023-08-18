# Functions

# Format
## Inc(FileName,SheetName,Saved, Month)
## Trs(FileName,SheetName,Saved, Month)

# !run
Inc("AWPJun23.xlsx", "AWP1","Indicators_Results", "Jun")
Trs("AWPJun23.xlsx", "AWP2","TrainSupport_Results", "Jun")

Pool("Indicators_Results.csv","TrainSupport_Results.csv")
