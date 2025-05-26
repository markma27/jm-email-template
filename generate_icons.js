const { createCanvas } = require('canvas');
const fs = require('fs');

// Icon sizes needed
const sizes = [16, 48, 128];

function createIcon(size) {
    // Create canvas
    const canvas = createCanvas(size, size);
    const ctx = canvas.getContext('2d');
    
    // Background
    ctx.fillStyle = '#0078d4';
    ctx.fillRect(0, 0, size, size);
    
    // Email icon (simplified)
    ctx.fillStyle = 'white';
    const margin = size * 0.2;
    const width = size - (margin * 2);
    const height = width * 0.7;
    const x = margin;
    const y = (size - height) / 2;
    
    // Email body
    ctx.fillRect(x, y, width, height);
    
    // Email flap
    ctx.fillStyle = '#106ebe';
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + width/2, y + height/3);
    ctx.lineTo(x + width, y);
    ctx.closePath();
    ctx.fill();
    
    return canvas;
}

// Generate icons for each size
sizes.forEach(size => {
    const canvas = createIcon(size);
    const buffer = canvas.toBuffer('image/png');
    fs.writeFileSync(`icon${size}.png`, buffer);
    console.log(`Generated icon${size}.png`);
}); 