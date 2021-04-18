# โปรแกรมจัดการข้อมูลสถิติการฝึกงานของนักศึกษาคณะวิทยาศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลธัญบุรี ย้อนหลัง3ปี
เป็นฐานข้อมูลสถานประกอบการและโปรแกรมจัดการฐานข้อมูลได้ไปในตัว
ทดลองดาวน์โหลด excel มารันได้เลยอย่าลืมอนุญาติให้รัน VBA code ("คลิกเปิดใช้งานเนื้อหา") ด้วย ปลอดภัยไม่ใช่ไวรัสครับ 
<p align="center">
  <img src="https://firebasestorage.googleapis.com/v0/b/sorawitwebsite.appspot.com/o/VBAexcel111.png?alt=media&token=0eb0d9f7-cd6e-47e0-bc71-a2679bdbfac0" width=70%" title="hover text">
</p>
 เนื่องจากคณะวิทยาศาสตร์ มหาวิทยาลัยเทคโนโลยีราชมงคลธัญบุรี มีข้อมูลสถิติการฝึกงานเป็นไฟล์ excel นี้และได้มอบหมายให้มขณะที่กำลังฝึกงานในมหาวิทยาลัยสร้างกราฟสถิติจากข้อมูลนี้ขึ้นมา ผมจึงมีไอเดียสร้างโปรแกรมจัดการข้อมูลนี้จากไฟล์ excel ซะเลย ซึ่งได้ใช้ VBA script ใน excel ในการสร้างโปรแกรม
<p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel1.png" width=70%" title="hover text">
</p>
 โปรแกรมประกอบได้วย หน้า(sheet) "data","section","original" คือต้องบอกก่อนว่าไฟล์นี้มีปัญหา ก่อนที่ผมรับช่วงต่อ ข้อมูลนักศึกษาดันมี 2 หน้า(sheet) แยกกันเป็น หน้ารวมสถิตินักศึกษาฝึกงานทั้ง 7 สาขา และ หน้าสถิติฝึกงานของนักศึกษาสาขาฟิสิกส์ตามลำดับ ทางมหาลัยจึงมอบหมายให้ผมทำหน้าสถิตินักศึกษาฝึกงานทั้ง 7 สาขาไปด้วย ผมจึงมีไอเดีย(อีกแล้ว) ให้มันสร้างหน้าข้อมูลอัตโนมัติของแต่ละสาขาไปเลย</br>
  <b> การจัดการข้อมูล </b>
 การแก้ไขจำนวนนักศึกษา เนื่องจากมีข้อมูลสถานประกอบการจำนวนมากเพื่อความสะดวกในการกรอกข้อมูล ผมจึงสร้างปุ่มค้นหาสถานประกอบการขึ้นมา
 <p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel4.png" width=70%" title="hover text">
</p>
กรอกชื่อสถานประกอบการ(ไม่จำเป็นต้องกรอกชื่อเต็ม)แล้วกดค้นหาโปรแกรมจะกรองข้อมูลชื่อที่คล้ายกันออกมา
<p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel5.png" width=70%" title="hover text">
</p>
แก้ไขข้อมูลได้ตามใจชอบ ตามปีที่ใช้งาน หรือ จำนวนนักศึกษาแต่ละสาขา ตามชื่อย่อดังนี้ ( B= ชีวะ, CH= เคมี, PH= ฟิสิกส์, ST= สถิติ, CS= วิทยาการคอม, CT= เทคโนคอม ,M= คณิต )
สร้างหน้าข้อมูลสรุปสถิติของแต่ละสาขา ในเคสนี้คือ สาขาคณิตสาสตร์ Math ให้คลิกเลือกสาขา
<p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel6.png" width=70%" title="hover text">
</p>
เมื่อตกลงโปรแกรมจะสร้างหน้าสรุปข้อมูลสาขา Math อัตโนมัติใช้เวลาสักครู่ ช้าเร็วขึ้นอยู่กับปริมาณข้อมูลและความแรงของคอมฯ 
<p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel7.png" width=70%" title="hover text">
</p>    
สามารถคลิกสร้างกราฟ ก็จะสร้างกราฟให้อัตโนมัติ
<p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel8.png" width=70%" title="hover text">
</p> 
 การอัปเดตข้อมูล เมื่อปีไหนที่นักศึกษาได้สถานประกอบการที่ไม่เคยมีอยู่ในฐานข้อมูลสามารถอัปเดตสถานประกอบการได้
 <p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel2.png" width=70%" title="hover text">
</p>  
จะเพื่มชื่อสถานประกอบการใหม่ในแถวล่าสุด
 <p align="center">
  <img src="https://raw.githubusercontent.com/thekem0/Apprenticeships_statistics/main/IMG/vbaexcel3.png" width=70%" title="hover text">
</p> 
