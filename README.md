
# ExcelConverter for Unity
An editor tool that allows you convert excel(xlsx/xls) into json/bson files.

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/ExcelConverter.png" height="50%" width="50%"/>

## Quick Installnation:
### Required for installation:
- Unity **2020.3.11** or higher
- Unity version below **2020.3.11** needs install [Json.NET](https://github.com/jilleJr/Newtonsoft.Json-for-Unity)

### Install via git URL
Add url below to Package Manager.
``` 
https://github.com/JohnsonChenz/ExcelConverter.git?path=Assets/Plugins/ExcelConverter 
```

## Features:
- Convert excel into json/bson
- Deploy export config of excels via one single json file.
- Quick create/setup json config by using json config generator.

## How to use:
"Plugins -> ExcelConverter" to open it.
1. Choose source file folder.
2. Choose output file folder.
3. Choose an json config file. (About how to create json config, please look below:
4. Choose export option (json/bson/both).
5. Click Export button to generate file.
- **Example of Excel & Json Config file is provided in project**.

## Setup you own json config with json config generator

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/JsonConfigGenerator.png" height="50%" width="50%"/>

- The way your excel file being converted is based on the Mainkey/Subkey rule you like in the excel,so,you will need to create a json config file to determine how your excel file will be converted into json/bson data.  
- In **ExcelConverter window**, click the button **Json Config Generator** to open the generator. 
- Create and setup single/multiple json config based on your needs.
- Save set-up json config to json file somewhere you like,then browse the path of it in ExcelConverter.
- When converting,ExcelConverter will load json config file from your browsed path to convert excel file into json/bson based on the setting of the json config.

## Simple explanation of Json Config Class
**MainKeyType (enum)**
- Determine mainkey of converted data is uppercase or not.

**SubKeyType (enum)**
- Determine subkey of converted data is uppercase or not.

**MainKeyColumn (int)**
- Determine which column of excel that will added as mainkey,if mainkey data is valid,converted json data will be **json object(map)**
- Set it as 0 when you don't wanna add mainkey to your converted data,converted json data will be **json array(array)**.
- Extra : When you mainkey data is composed of **multiple columns** in excel sheet,for example,when your prefered mainkey data is made up of **column 1** + **column 2** in excel sheet,set the value to 2.

**SubKeyRow (int)**
- Determine which row of excel will be added as subkey, if subkey data is valid, actual data of excel will package with **json object(map)**
- Set it as 0 when you don't wanna add subkey to your converted data, actual data of excel will package with **json array(map)**

**FirstDataRow (int)**
- Determine the which row for ExcelConverter to start reading as actual data when converting.

**Datalist (string array)**
- List of sheet name that will apply settings above for converting.

## Converted result showcase
Excel sheet :  

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/ExcelSheet.png" height="80%" width="80%"/>

### Mainkey + Subkey :

Config :

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/WithMainKeyAndSubKey.png" height="50%" width="50%"/>

Result :
```json
{
	"export_type": "JObject",
	"data": {
		"JOHN": {
			"SUBKEY_GENDER": "Male",
			"SUBKEY_AGE": 15,
			"SUBKEY_HEIGHT": 170,
			"SUBKEY_WEIGHT": "65kg",
			"SUBKEY_ARRAY": [
				"Str_1",
				"Str_2",
				"Str_3"
			]
		},
		"MARRY": {
			"SUBKEY_GENDER": "Female",
			"SUBKEY_AGE": 9,
			"SUBKEY_HEIGHT": 156,
			"SUBKEY_WEIGHT": "45kg",
			"SUBKEY_ARRAY": [
				1,
				2,
				3
			]
		},
		"KEN": {
			"SUBKEY_GENDER": "Male",
			"SUBKEY_AGE": 23,
			"SUBKEY_HEIGHT": 182,
			"SUBKEY_WEIGHT": "70kg",
			"SUBKEY_ARRAY": [
				true,
				false,
				true
			]
		}
	}
}
```

### Mainkey only :

Config :

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/WithMainKeyOnly.png" height="50%" width="50%"/>

Result :
```json
{
	"export_type": "JObject",
	"data": {
		"JOHN": [
			"Male",
			15,
			170,
			"65kg",
			[
				"Str_1",
				"Str_2",
				"Str_3"
			],
			"whatever",
			"whatever"
		],
		"MARRY": [
			"Female",
			9,
			156,
			"45kg",
			[
				1,
				2,
				3
			],
			"whatever",
			"whatever"
		],
		"KEN": [
			"Male",
			23,
			182,
			"70kg",
			[
				true,
				false,
				true
			],
			"whatever",
			"whatever"
		]
	}
}
```

### Subkey only :

Config :

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/WithSubkeyOnly.png" height="50%" width="50%"/>

Result :
```json
{
	"export_type": "JArray",
	"data": [
		{
			"MAINKEYCOLUMN": "John",
			"SUBKEY_GENDER": "Male",
			"SUBKEY_AGE": 15,
			"SUBKEY_HEIGHT": 170,
			"SUBKEY_WEIGHT": "65kg",
			"SUBKEY_ARRAY": [
				"Str_1",
				"Str_2",
				"Str_3"
			]
		},
		{
			"MAINKEYCOLUMN": "Marry",
			"SUBKEY_GENDER": "Female",
			"SUBKEY_AGE": 9,
			"SUBKEY_HEIGHT": 156,
			"SUBKEY_WEIGHT": "45kg",
			"SUBKEY_ARRAY": [
				1,
				2,
				3
			]
		},
		{
			"MAINKEYCOLUMN": "Ken",
			"SUBKEY_GENDER": "Male",
			"SUBKEY_AGE": 23,
			"SUBKEY_HEIGHT": 182,
			"SUBKEY_WEIGHT": "70kg",
			"SUBKEY_ARRAY": [
				true,
				false,
				true
			]
		}
	]
}
```

### Double mainkey :

Config :

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/DoubleMainKey.png" height="50%" width="50%"/>

Result :
```json
{
	"export_type": "JObject",
	"data": {
		"JOHNMALE": {
			"SUBKEY_AGE": 15,
			"SUBKEY_HEIGHT": 170,
			"SUBKEY_WEIGHT": "65kg",
			"SUBKEY_ARRAY": [
				"Str_1",
				"Str_2",
				"Str_3"
			]
		},
		"MARRYFEMALE": {
			"SUBKEY_AGE": 9,
			"SUBKEY_HEIGHT": 156,
			"SUBKEY_WEIGHT": "45kg",
			"SUBKEY_ARRAY": [
				1,
				2,
				3
			]
		},
		"KENMALE": {
			"SUBKEY_AGE": 23,
			"SUBKEY_HEIGHT": 182,
			"SUBKEY_WEIGHT": "70kg",
			"SUBKEY_ARRAY": [
				true,
				false,
				true
			]
		}
	}
}
```

### No key :

Config :

<img src="https://github.com/JohnsonChenz/ExcelConverter/blob/master/Docs/WithoutMainKeyAndSubKey.png" height="50%" width="50%"/>

Result :
```json
{
	"export_type": "JArray",
	"data": [
		[
			"John",
			"Male",
			15,
			170,
			"65kg",
			[
				"Str_1",
				"Str_2",
				"Str_3"
			],
			"whatever",
			"whatever"
		],
		[
			"Marry",
			"Female",
			9,
			156,
			"45kg",
			[
				1,
				2,
				3
			],
			"whatever",
			"whatever"
		],
		[
			"Ken",
			"Male",
			23,
			182,
			"70kg",
			[
				true,
				false,
				true
			],
			"whatever",
			"whatever"
		]
	]
}
```

## Converting rules about excel
1. If a certain subkey field's name contains symbol \* or filled as "empty",it's column will be totally ignored read into converted data.
2. Actual data field that filled in like [element1,element2,element3.....] will be convert into data as **json array**.

## License
This library is under the MIT License.
