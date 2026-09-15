Add-Type -AssemblyName System.Drawing
$imgPath = "C:\Users\BINYAMEEN\.gemini\antigravity-ide\brain\2e7ad517-12aa-427a-b1cb-e4b713d270a2\media__1786785436024.png"
$img = [System.Drawing.Image]::FromFile($imgPath)
$chunks = 8
$chunkH = [math]::Floor($img.Height / $chunks)
$scale = 3

for ($i = 0; $i -lt $chunks; $i++) {
    $outW = $img.Width * $scale
    $outH = $chunkH * $scale
    $crop = New-Object System.Drawing.Bitmap($outW, $outH)
    $g = [System.Drawing.Graphics]::FromImage($crop)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality

    $srcRect = New-Object System.Drawing.Rectangle(0, ($i * $chunkH), $img.Width, $chunkH)
    $destRect = New-Object System.Drawing.Rectangle(0, 0, $outW, $outH)
    $g.DrawImage($img, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
    
    $outPath = "C:\Users\BINYAMEEN\.gemini\antigravity-ide\brain\2e7ad517-12aa-427a-b1cb-e4b713d270a2\crop_scaled_" + $i + ".png"
    $crop.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $crop.Dispose()
}
$img.Dispose()
Write-Host "Done cropping scaled chunks"
