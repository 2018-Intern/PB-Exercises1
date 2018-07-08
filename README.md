# PB-Exercises1

## 前置動作

> Open Windows 

* 檔案位置: Library/Application(exercises.pbl)

```
Open(w_main) // 打開主畫面
```

> Database Connect Syntax

* 檔案位置: w_main的Open Event

```
// Profile sun (可至DB Profile 內 copy)
SQLCA.DBMS = "SNC SQL Native Client(OLE DB)"
SQLCA.LogPass = "123456" 
SQLCA.ServerName = "sun-pc\sqlexpress"
SQLCA.LogId = "sa"
SQLCA.AutoCommit = False
SQLCA.DBParm = "Database='sun_test',Provider='SQLNCLI10'"

//Conncet 
CONNECT USING SQLCA; //開始連接SQLCA

//訊息碼判斷 0 => 執行成功。 | 100 => 成功但無資料。 | <0 => 有錯誤。
if SQLCA.SQLCode = 0 then
	MessageBox("","Connect sucessful !") //成功
else 
	MessageBox("","Connect error !")  //有錯誤
end if

```

## SQL Connect+Retrieve+Insertrow+Update

![展示](https://github.com/2018-Intern/PB-Exercises1/blob/master/exercise1.png)

* SQL Connect 資料庫連線
```
//Connect Button Clicked Event
  dw_1.settransobject(SQLCA) 
```

* SQL Retrieve 讀取資料
```
//Retrieve Button Clicked Event
  dw_1.retrieve()
```

* SQL Insertrow 插入一列
```
//Insertrow Button Clicked Event
  dw_1.insertrow(0) // 參數為0，代表在當前Datawindow的最後一行插入一空行
```

* SQL Update 更新資料
```
//Update Button Clicked Event
  dw_1.update()
  
  if dw_1.update() = 1 then
    Commit USING SQLCA;
  else 
    ROLLBACK USING SQLCA;
  end if
```
