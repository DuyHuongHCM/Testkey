# Testkey

Tool kiểm tra nhanh tình trạng bản quyền Windows và Office.

## Cách dùng

Mở PowerShell với quyền **Run as Administrator** và chạy:

```powershell
powershell -ExecutionPolicy Bypass -File .\check-license.ps1
```

## Kết quả

- **Windows**: chạy `slmgr /xpr` để hiển thị trạng thái kích hoạt.
- **Office**: tìm `OSPP.VBS` và chạy `/dstatus` để hiển thị trạng thái bản quyền của Office.
