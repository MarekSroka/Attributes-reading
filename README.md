# Reading_data_project_TS
Reading data from *.xlsx files

Skrypt przegląda wszystkie pliki Excel w wybranym folderze - folder podawany jako parametr wejściowy wpisywany w konsoli.

Bierze tylko wiersze, gdzie objectType nie jest pusty.

Pobiera wszystkie wartości z kolumn o nazwie zawierającej "attributeList.attribute.name...".

pobiera wszystkie wartości z kolumn, w kórych określone są 'Attribute Value' (np. "attributeList.attribute.string").

Jeśli po kolumnie zawierającej "attributeList.attribute.name..." jest kolumna attributeList.attribute.<typ_kolumny>, to pobiera wartość z tej kolumny, jeśli brak wartości, to wypisuje "".

Nazwa pliku wyjściowego to nazwa folderu, np. Component.xlsx.
