# EngWordBank
自製簡單英文單字庫APP

程式語言：`python`

主要匯入模組：`Tkinker`,`openpyxl`

操作說明事項：

1.必須把`.xlsx`檔和`exe.`檔放在同一個目錄下，才能正常讀取。(第一次執行exe.檔後，可直接建立單字，`wordbank.xlsx`檔會自動產生)

2.開啟exe.進入使用者介面後，一共有三個分頁：
  
  (一)新增單字：以下圖為例，假設在電影《末日之戰》上看到doomsday這單字，將查詢後的中文意思輸入後，按下新增鈕，就會在`wordbank.xlsx`中創建名為YouTube的工作表。
  ![圖片](https://user-images.githubusercontent.com/103128273/183119425-e89b2ed5-37f1-459d-afaa-81f93351977c.png)

  (二)查詢單字：查詢時輸入單字，按下查詢，就會跑出結果囉!
  ![圖片](https://user-images.githubusercontent.com/103128273/183119506-8a1908e1-b76b-4365-b0c0-441db8208cd2.png)

  
  (三)工作表名稱庫：自己新增的單字出處之名稱表。當要在"同一個"工作表中新增單字時，需要參考這個表，以輸入相同的工作表名稱。(按下更新即會跑出內容)
  ![圖片](https://user-images.githubusercontent.com/103128273/183120333-6ca524e5-1409-4fb9-b694-943490ef6893.png)




English word bank APP designed by myself

Must put the `wordbank.xlsx` file(auto generated) which is including English words in the current work folder or directory.

製作原因：
在學習英文單字的過程中，常常會在不同地方遇到不懂的英文單字，上網查詢後過沒多久又會忘記。因此我想製作一個專屬於個人的英文單字庫，可以將查訊後的單字和其翻譯後的意思，依自己喜好分門別類地歸納，下次當遇到自己可能看過的單字時，就能在此APP單字庫裡查詢自己曾經在哪裡看過這單字，出處是哪裡，以加強對這個單字的印象。
