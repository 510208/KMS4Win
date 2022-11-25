# KMS4Win

## 目錄
[💯 前言](#-前言)

[📢 注意](#-注意)

[🔑 密碼](#-密碼)

[😍 關於我](#-關於我)

[🍿 改密碼](#-改密碼)
___

### 💯 前言

此軟體可以啟用Windows與Office，完全免費及開源！

[![GitHub stars](https://img.shields.io/github/stars/510208/KMS4Win/?color=brightgreen&style=for-the-badge)](https://github.com/510208/KMS4Win/)
[![](https://img.shields.io/badge/Blog-510208's%20Blog-brightgreen?style=for-the-badge&logo=appveyor)](https://sam0616.pixnet.net)
[![MIT license](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=for-the-badge&logo=appveyor)]

____

### 📢 注意

此軟體並不會鎖定你的電腦，而是開啟一視窗定位於最上層；並且如果你忘記密碼，可以裝置強制重啟以恢復(當然，前提是你並未將軟體加入自啟動目錄)

此軟體開源託管在GitHub上，完全免費。如果你付費到此下載，請給他檢舉！

我並不想拿此軟體作為營利用途，只是寫爽的(XD)

___

### 🔑 密碼

此軟體預設密碼為
***510208***
，如需自行更改密碼請
[下載VB6.0編譯工具](https://drive.google.com/file/d/1dqhy93UtBeTGgbp09Jk6wN_A8kj5280H/view?usp=sharing)
感謝使用。

___

### 😍 關於我

[![YouTube](https://img.shields.io/badge/YouTube-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-red?style=for-the-badge&logo=appveyor)](https://www.youtube.com/channel/UC6orwHdQNVzwHsA6M7HYD9g/videos?view=0&sort=p&shelf_id=0)
[![Blog](https://img.shields.io/badge/Pixnet-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-blue?style=for-the-badge)](https://sam0616.pixnet.net)
[![WordPress](https://img.shields.io/badge/WordPress-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-yellowgreen?style=for-the-badge&logo=appveyor)](https://510208.nde.tw)

喜歡這支程式，可以訂閱點讚開小鈴鐺、部落格新增到書籤等...^^

___

### 🍿 改密碼

程式中有段代碼如下：

```VB.NET
If Text1.Text = "510208(換成你要的密碼)" Then
    MsgBox "密碼正確，按下確定以解鎖", 64
    End
Else
    MsgBox "密碼錯誤", 16
    Label3.Caption = "密碼提示：網名"
End If
End Sub
```

將510208處換成你要的密碼，譬如我想要密碼改為"52penguinbrother"，如下修改：

```VB.NET
If Text1.Text = "52penguinbrother" Then
    MsgBox "密碼正確，按下確定以解鎖", 64
    End
Else
    MsgBox "密碼錯誤", 16
    Label3.Caption = "密碼提示：網名"
End If
End Sub
```

接下來將其編譯為exe檔案：

![image](https://img.onl/4I0erl)

