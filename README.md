# [Excel VBA] appFormToJson : Excel定型帳票から入力情報をJSON形式で抽出するツール

## overveiw :

- Excelで作成した定型フォーマットの申請書、等から、入力情報を抽出する補助ツール(VBA Excel)。
- 定型フォーマットから取得する値を、レイアウト(XXX.xlsx)で定義し、その定義をもとに、抽出定義json(XXX.xlsx.json)を生成する。抽出定義jsonをもとに、入力情報を抽出する。  
    - 抽出定義jsonは、マクロにより生成可能。

### 機能 :
- Excelに入力された定型フォーマットを、指定フォルダ（INPUT_FOLDER)に配置する。  
    + レイアウトより、抽出定義(XXX.xlsx.json)を生成し、`DEF_FOLDER`に配置する。  

- メニューより、以下の機能を利用する。
    + 【Output JSON data from a Sheet】抽出定義をもとに定型フォームからJSON出力する。（サンプル）  
    + 【Dump a def json】レイアウトより、抽出定義(XXX.xlsx.json)を出力する。

## Installation :

- GitHubより、Cloneする。  
    + https://github.com/sakai-memoru/appFormToJson  

- 参照設定が必要。  

![referto](https://gyazo.com/7d30f2387e7818067fd7596a82e507e9.png) 


## Usage :
- アプリは以下。
    - アプリ本体  ：appFormToJson.xlsm  
        + Batch   : FormToJsonMain.bas  
            + FormToJsonModule.bas  
                - GetValue 
                - DumpSimpleJson - GetDef 
  - アプリconfig：config.json  
  - 抽出定義form：  
      + defs/RequestSheet.xlsm 申請フォーマット（サンプル）  

- appSpecDef.xlsmを開く。Menuより起動する。 

![menu](https://gyazo.com/e79bbe22eb2f614940404b5bd1b62a7b.png)  



### 初期コンフィグ設定 :
   
```
{
    "BASE_FOLDER": "",
    "INPUT_FOLDER": "input",
    "OUTPUT_FOLDER": "output",
    "BACKUP_FOLDER": "input/backup",
    "DEF_FOLDER": "defs",
    "RequestSheet": {
        "SHEET_TYPE": "SHEETFORM",
        "INPUT_LIKE": "Request*.xlsx",
        "DEF_WORKBOOK_NAME": "RequestSheet.xlsx",
        "DEF_SHEET_NAME": "Sheet",
        "DEF_NAME_PARAM": "layout",
        "DEF_FILE": "RequestSheet.xlsx.json",
        "MACRO_GET_METHOD": "GetValue",
        "MACRO_DUMP_METHOD": "DumpSimpleJson"
    },
    "CONTROL_PREFIX": "__",
    "SOURCE_FROM": "_source",
    "APP_NAME" : "appFormToJson"
}
```

### Environment

![env](https://gyazo.com/b0c2ce3be04ba8f4e8b29044ed2e425a.png)


## Execution sample

- レイアウト定義 : forms/_mapTableDesign.xlsm  
![map](https://gyazo.com/4a411aa38a51b24fdf1a6484938805ae.png)  

- 出力例  
```
{
    "applicantNameKana": "ニホンバシ　ジロウ",
    "applicantSignImage": "",
    "applicantName": "日本橋　二郎",
    "applicantBirthDate": "1985年10月01日",
    "applicantAddress": "東京都中央区日本橋１－１－２０２０",
    "applicationDate": "2020年6月17日",
    "applicantTel": "080-9876-5432",
    "reciever1": "日本橋二郎",
    "revalent1": "本人",
    "recieverBirthDate1": "1985年10月01日",
    "recieverFlag1": "要",
    "reciever2": "日本橋花子",
    "revalent2": "長女",
    "recieverBirthDate2": "2005年01月10日",
      :
      :
      :
      :
    "recieverFlag6": "",
    "accountNameKana": "ニホンバシジロウ",
    "bankName": "住菱銀行",
    "branchName": "日本橋支店",
    "accountClass": "普通",
    "bankCode": "9999",
    "branchCode": "1010",
    "accountId": "12345678",
    "accountPBNameKana": "",
    "accountPBCode": "",
    "accountPBId": "",
    "_source": "'RequestSheet_200616.xlsx'!Sheet",
    "_source_date": "2020-06-16T07:45:25.000Z",
    "_created": "2020-06-21T12:14:56.000Z",
    "_id": "D084C1E6-8850-7E55-9A72-8C9ABF49A5D2"
}
```

## application I/F

```vb
'''' **********************************************
'' @file FormToJsonMain.bas
'' @parent appFormToJson.xlsm
''

Public Function Batch( _
        ByVal datatype As String, _
        Optional ByVal dumpOn As Variant = False, _
        Optional ByVal moveOn As Variant = False _
    ) As Variant
'''' **********************************************
'''' @function batch
'''' @param datatype {String} 処理データタイプ
''''        config.jsonのキー "RequestSheet"
'''' @param dumpOn  {Variant<boolean>}
''''            dump def.json
'''' @param moveOn  {Variant<boolean>}
''''            Inputファイル移動flag
''
```  

## note :
- 落ち着いたら、もう少し記述を追加します。  

## reference :

- 以下の外部ライブラリを使用しています。  
  + VBA-JSON : JsonConverter.bas  
    - https://github.com/VBA-tools/VBA-JSON  
  + MiniTemplator  
    - https://www.source-code.biz/MiniTemplator/  

// --- end of README.md