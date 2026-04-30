Lista de validaciones

-Referencias tienen que limpiarse con =_xlfn.REGEXEXTRACT(F6,"\w?\w-J-\d{4}LI\d{2}")
-Si no existen referencias tendremos que usar BU... Por lo regular empiezan con M-m y seguido de un numero
-Si vemos que en Part Number existen cosas como TAPA PLASTICA, CHAROLA, BASE PLASTICA,etc. Se marcara su BU como Miscelaneus


-En SEA utilizamos algunas columnas para el mes de abril como BU, ITEM CODE, CONTAINER, TOTAL GROSS WEIGHT
-Habra momentos que el BU sera llamado 'Capex' ya que su item code dira CAPEX o solo dira letras sin numeros, aqui te doy un ejemplo CAPEX-08 o TR-TOOLING
