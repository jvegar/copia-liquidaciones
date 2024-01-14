# Global Variables
$Url = "https://intercorpretail.sharepoint.com/sites/AppsCorporativas/er/"
$BibliotecaLiquidacionesID = "099a87e4-64fc-45ee-b787-b422e1191b25"
$ListaLiquidacionDetalleID = "d5798787-02cc-490d-b53b-b6a475106629"
$ListaLiquidacionCabeceraID = "aa404da0-6a55-4288-9cb1-a1c08ae46ade"
$OneDriveUrl = "https://intercorpretail-my.sharepoint.com/personal/ir-entregas_rendir_corp_intercorpretail_pe"

# Connect to SP Site
Connect-PnPOnline $url -Interactive

# Get items from libraries and lists
$bibliotecaLiquidacionesItems = (Get-PnPListItem -List $BibliotecaLiquidacionesID -Fields "Detalle" -PageSize 5000).FieldValues
$listaLiquidacionesDetalleItems = (Get-PnPListItem -List $ListaLiquidacionDetalleID -Fields "ID","Liquidacion" -PageSize 5000).FieldValues
$listaLiquidacionCabeceraItems = (Get-PnPListItem -List $ListaLiquidacionCabeceraID -Fields "ID","Sociedad", "AnioSAP", "Solicitud" -PageSize 5000).FieldValues

# Log Function
function Log-Message([string] $message, [string] $type){
    Add-Content -Path ".\Log$($type).txt" $message
}

# Function to get Details by ID
function Get-Liquidaciones-Detalles ($liquidacionItem, $listaLiquidacionesDetalleItems, $listaLiquidacionCabeceraItems) {
    $liquidacionID = $null
    $liquidacionFileRef = $null
    foreach($detalleItem In $listaLiquidacionesDetalleItems) {
        Log-Message 
        if($detalleItem.ID -eq $liquidacionItem.Detalle) {
            $liquidacionID = $detalleItem.Liquidacion.LookupId
            $liquidacionFileRef = $liquidacionItem.FileRef
            if($null -ne $liquidacionID) {
                foreach($cabeceraItem In $listaLiquidacionCabeceraItems) {
                    if($cabeceraItem.ID -eq $liquidacionID) {
                        return [PSCustomObject]@{
                            Sociedad = $cabeceraItem.Sociedad
                            Anio = $cabeceraItem.AnioSAP
                            Solicitud = $cabeceraItem.Solicitud.LookupValue
                            FileRef = $liquidacionFileRef
                        }
                    }
                }
            }
        }
    }
    return $null
}

# Get liquidaciones details
$count = 0
$liquidaciones = New-Object System.Collections.ArrayList
ForEach($item In $bibliotecaLiquidacionesItems){
    Write-Host "Iteration (Detalle: $($item.Detalle)) #$count from $($bibliotecaLiquidacionesItems.Count) :"
    if ([int]$item.Detalle -gt 100000) {
        $liquidacion = Get-Liquidaciones-Detalles $item $listaLiquidacionesDetalleItems $listaLiquidacionCabeceraItems
        $liquidaciones.Add($liquidacion)
    }
    $count= $count+1
}