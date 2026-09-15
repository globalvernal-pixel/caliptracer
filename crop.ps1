Add-Type -AssemblyName System.Drawing
$imgPath = "C:\Users\BINYAMEEN\.gemini\antigravity-ide\brain\2e7ad517-12aa-427a-b1cb-e4b713d270a2\media__1786785436024.png"
$img = [System.Drawing.Image]::FromFile($imgPath)
Write-Host "Width:" $img.Width "Height:" $img.Height
$chunks = 8
$chunkH = [math]::Floor($img.Height / $chunks)

for ($i = 0; $i -lt $chunks; $i++) {
    $crop = New-Object System.Drawing.Bitmap($img.Width, $chunkH)
    $g = [System.Drawing.Graphics]::FromImage($crop)
    $srcRect = New-Object System.Drawing.Rectangle(0, ($i * $chunkH), $img.Width, $chunkH)
    $destRect = New-Object System.Drawing.Rectangle(0, 0, $img.Width, $chunkH)
    $g.DrawImage($img, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
    $outPath = "C:\Users\BINYAMEEN\.gemini\antigravity-ide\brain\2e7ad517-12aa-427a-b1cb-e4b713d270a2\crop_" + $i + ".png"
    $crop.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $crop.Dispose()
}
$img.Dispose()
Write-Host "Done cropping 8 chunks"
