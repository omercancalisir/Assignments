-Programların çalışabilmesi için CaseWare IDEA programının açık olması gerekmektedir.

-Assignment1 programında kullanmak için IDEA özelliği olan Visual Connector kullanarak; CUSTNO, CREDIT_LIM ve
AMOUNT isimli sütunların olduğu CMF-BT.IMD adlı veri tabanını oluşturdum.

-Assignment2 programında condition kısmında örneğin CREDIT_LIM değerleri 4000 değerinden küçükleri istiyorsanız
"< 4000" yazmanız yeterlidir. Aynı şekilde CREDIT_LIM değerleri 20000'e eşit olanları istediğinizdeyse "= 20000" 
yazmanı yeterlidir. Örneğin COUNTRY sütununu seçtiyseniz Condition kısmına SOUTH AFRICA yazmanız yeterlidir.

- .IMD veri tabanı dosyalarının Pandas dataframe e dönüştürülmesi için IDEALib modülü kullanıldı. Verilerin
gösterimi için ise Tkinter ve pandastable kütüphanesi kullanıldı.
