# Script Nhắc nhở cập nhật daily_notes.md hàng ngày
$dailyNotesPath = "d:\AI AGENT THUY\AI branding\scratch\daily_notes.md"
$title = "Nhắc nhở Daily Notes - Hệ thống PBAN"
$message = "Chị Thúy ơi, hãy dành 3 phút ghi lại câu chuyện thực tế hôm nay tại Eureka vào file daily_notes.md nhé!"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
