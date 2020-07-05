const puppeteer = require('puppeteer');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

(async () => {
    // if(fs.existsSync('image-cpature-export-ppt.pptx')){
    //     fs.unlinkSync('image-cpature-export-ppt.pptx')
    //     console.log('deleted existing ppt')
    // }
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.setViewport({
        width: 1920,
        height: 1080,
        deviceScaleFactor: 1,
      });
    await page.goto('https://en.wikipedia.org/wiki/Main_Page');
    await page.waitForSelector('#mp-left');
    let element = await page.$('#mp-left');
    await element.screenshot({path: '1.png'}); 
    element = await page.$('#mp-right');
    await element.screenshot({path: '2.png'});
    await page.waitForSelector('#mp-lower');
    element = await page.$('#mp-lower');
    await element.screenshot({path: '3.png'}); 
    await browser.close();
    // define master slides
    let pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.defineSlideMaster({
        title: "MASTER_SLIDE",
        bkgd: "#7b94d1",
        objects: [
            { rect: { x: 0.5, y: 0.2, w: "95%", h: 0.75, fill: "F1F1F1" } },
            { text: { text: "This is a powerpoint slide created with JavaScript", options: { x: 0.5, y: 0.2, w: "95%", h: 0.75 } } },
            { image: { x: 11.3, y: 6.8, w: 1.50, h: 0.50, path: "logo.jpg" } },
            { text: { text: "Slide: ", options: { x: 0.1, y: "90%", color: 'F1F1F1', w: "95%" , fontSize: 16} } }
        ],
        slideNumber: { x: 0.6, y: "89.5%", color:'F1F1F1', fontSize: 16 },
    });
    let slide1 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
    slide1.addText("Placeholder holding custom values here!", { x: 0.6, y: 1.5, w: 12, h: 5.25, color: "FFFFFF", fontSize: 24 },);
    for (let i=1; i<=3; i++) {
        let width = 12.0, height;
        if ((i==1 || i==2)) {
            height = 5.5;
        } else {
            height = 5.0;
        }
        let slide = pptx.addSlide({ masterName: "MASTER_SLIDE" })
        slide.addImage({path: `${i}.png`, x:0.5, y:1.2, w:width, h:height });
        slide.addNotes('This slide holds the image captured from wikipedia through puppeteer!');    
    }
    let slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
    slide3.addText("Thank you!", { x: 3.5, y: 3.0, w: 7.5, h: 1.0, color: "FFFFFF", fontSize: 48 });
    pptx.writeFile('image-capture-export-ppt.pptx')
    .then(fileName => {
        console.log(`created file: ${fileName}`);
    });
})();