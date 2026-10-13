// Power Query (M) — Limpieza de la exportación de tickets
// Uso: Datos > Obtener datos > Desde tabla/rango (sobre Datos_Crudos) > Editor avanzado > pegar este código
let
    Origen = Excel.CurrentWorkbook(){[Name="tblCrudos"]}[Content],

    // 1. Quitar filas totalmente vacías
    SinVacias = Table.SelectRows(Origen, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),

    // 2. Quitar duplicados exactos
    SinDuplicados = Table.Distinct(SinVacias),

    // 3. Recortar espacios y poner en Mayúscula Cada Palabra
    ColsTexto = {"Area_Solicitante","Categoria","Prioridad","Tecnico","Estado"},
    Recortado = Table.TransformColumns(SinDuplicados, List.Transform(ColsTexto, each {_, (t) => Text.Proper(Text.Trim(Text.From(t))), type text})),

    // 4. Unificar variantes (sinónimos y errores de digitación)
    Areas = Table.ReplaceValue(Recortado, each [Area_Solicitante], each
        if List.Contains({"Operacion"}, [Area_Solicitante]) then "Operaciones"
        else if [Area_Solicitante] = "Rrhh" then "Recursos Humanos"
        else if [Area_Solicitante] = "Ventas" then "Comercial"
        else if [Area_Solicitante] = "Logistica" then "Logística"
        else if [Area_Solicitante] = "Gerencia General" then "Gerencia"
        else [Area_Solicitante], Replacer.ReplaceValue, {"Area_Solicitante"}),
    Categorias = Table.ReplaceValue(Areas, each [Categoria], each
        if [Categoria] = "Red" then "Redes"
        else if List.Contains({"Accesos","Accesos Y Cuentas"}, [Categoria]) then "Accesos y cuentas"
        else if [Categoria] = "Impresora" then "Impresoras"
        else if [Categoria] = "E-Mail" then "Correo"
        else if [Categoria] = "Sofware" then "Software"
        else if [Categoria] = "Hardwre" then "Hardware"
        else [Categoria], Replacer.ReplaceValue, {"Categoria"}),
    Prioridades = Table.ReplaceValue(Categorias, "Critica", "Crítica", Replacer.ReplaceValue, {"Prioridad"}),
    Tecnicos = Table.TransformColumns(Prioridades, {{"Tecnico", each
        if Text.Upper(_) = "ANDRÉS ROJAS" then "Andrés Rojas"
        else if Text.Upper(_) = "LAURA GÓMEZ" then "Laura Gómez"
        else if Text.Upper(_) = "CAMILO PÉREZ" then "Camilo Pérez"
        else _, type text}}),

    // 5. Fechas: mezcla de fecha real y texto dd/mm/aaaa hh:mm -> datetime (configuración es-CO)
    Fechas = Table.TransformColumnTypes(Tecnicos, {{"Fecha_Apertura", type datetime}, {"Fecha_Cierre", type datetime}}, "es-CO"),

    // 6. Cierre anterior a la apertura -> estado "Revisar" y cierre nulo
    EstadoRevisar = Table.ReplaceValue(Fechas, each [Estado], each
        if [Fecha_Cierre] <> null and [Fecha_Cierre] < [Fecha_Apertura] then "Revisar" else [Estado],
        Replacer.ReplaceValue, {"Estado"}),
    CierreValido = Table.ReplaceValue(EstadoRevisar, each [Fecha_Cierre], each
        if [Estado] = "Revisar" then null else [Fecha_Cierre], Replacer.ReplaceValue, {"Fecha_Cierre"}),

    // 7. Satisfacción fuera de 1–5 -> nulo (no se imputa)
    Satisf = Table.TransformColumns(CierreValido, {{"Satisfaccion", each if _ <> null and _ >= 1 and _ <= 5 then _ else null, Int64.Type}}),

    // 8. Orden final
    Ordenado = Table.Sort(Satisf, {{"ID_Ticket", Order.Ascending}})
in
    Ordenado
