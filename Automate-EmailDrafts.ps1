$csv = Import-Csv "<List of Emails with DisplayName>"

$outlook = New-Object -comObject Outlook.Application
$displayName = @()

foreach($item in $csv)
{
    $Mail = $outlook.CreateItem(0)
    $Mail.Recipients.Add("$($item.DisplayName)")
    $Mail.Subject = "<Subject line>"
    $Mail.Attachments.Add("<Attachment file>")
    $Mail.HTMLBody = @"
    <p style='mso-layout-grid-align:none;text-autospace:none;font-size:11pt'><span style='mso-ascii-font-family:Calibri;mso-ascii-font-size:11.0pt;mso-hansi-font-family:Calibri;mso-hansi-font-size:11.0pt;mso-bidi-font-family:Calibri;mso-bidi-font-size:11.0pt;color:black;font-size:11.0pt'>
    <p style='font-size:11pt'>$($item.DisplayName),</p>
    <p style='font-size:11pt'>"Your message here"</p>
"@
    $Mail.Recipients.ResolveAll() | Out-Null
    $Mail.save()
}