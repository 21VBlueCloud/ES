# <CHNMachineFunctions
$HPSystemHealthTable = @{
    [UInt16]0="Unknown"
    [UInt16]5="OK"
    [UInt16]10="Degraded"
    [UInt16]15="Minor Failure"
    [UInt16]20="Major Failure"
    [UInt16]25="Critical Failure"
}

$HPRedundancyStatusTable=@{
    
    [UInt16]0="Unknown"
    [UInt16]2="Fully Redundant"
    [UInt16]3="Degraded Redundancy"
    [UInt16]4="Redundancy Lost"
    [UInt16]5="Overall Failure"
}

$HPOperationalStatusTable=@{
    [UInt16]0="Unknown"
    [UInt16]2="OK"
    [UInt16]3="Degraded"
    [UInt16]5="Predictive Failure"
    [UInt16]6="Error"
    [UInt16]10="Stop"
}

$HPSAOperationalStatusTable=@{
    [UInt16]0="Unknown"
    [UInt16]2="OK"
    [UInt16]5="Predictive Failure"
    [UInt16]6="Error"
}

$HPiLONICConditionTable=@{
    [UInt16]0="Unknown"
    [UInt16]2="OK"
    [UInt16]3="NIC disabled"
    [UInt16]4="NIC not in operation-alternate host NIC in use"
    [UInt16]5="NIC in operation but disconnected"
    [UInt16]6="Failed"
}
# CHNMachineFunctions>

# <CHNPerformance

# CHNPerformance>

# <CHNTenantFunctions
# CHNTenantFunctions>

# <CHNTopologyFunctions
# CHNTopologyFunctions>

# <MyUtilities
# MyUtilities>

# <SPSHealthChecker
# SPSHealthChecker>

# common helper
$CommonColor = @{
    Red = "#FF0000"
    Cyan = "#00FFFF"
    Blue = "#0000FF"
    purple = "#800080"
    Yellow = "#FFFF00"
    Lime = "#00FF00"
    Magenta = "#FF00FF"
    White = "#FFFFFF"
    Silver = "#C0C0C0"
    Gray = "#808080"
    Black = "#000000"
    Orange = "#FFA500"
    Brown = "#A52A2A"
    Maroom = "#800000"
    Green = "#008000"
    Olive = "#808000"
    Pink = "#FFC0CB"
    DarkBlue = "#00008B"
    DarkCyan = "#008B8B"
    DarkGray = "#A9A9A9"
    DarkGreen = "#006400"
    DarkMagenta = "#8B008B"
    DarkOrange = "#FF8C00"
    DarkRed = "#8B0000"
    DeepPink = "#FF1493"
    LightBlue = "#ADD8E6"
    LightCyan = "#E0FFFF"
    LightGray = "#D3D3D3"
    LightGreen = "#90EE90"
    LightYellow = "#FFFFE0"
    MediumBlue = "#0000CD"
    MediumPurple = "#9370DB"
    YellowGreen = "#9ACD32"
}