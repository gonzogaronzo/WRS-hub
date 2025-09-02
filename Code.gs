// Wilson Reading System\u00ae Slides Generator
// Google Apps Script implementation
// Creates Google Slides decks from word lists in Google Sheets

// --- Constants -------------------------------------------------------------
const WELDED_SOUNDS = [
  'all','am','an','ang','ank','ing','ink','ong','onk','ung','unk',
  'ild','ind','old','olt','ost'
];
const COLORS = {
  consonant: '#F5F5DC',
  vowel: '#F8B195',
  welded: '#9ACD32',
  affix: '#FFD700',
  card: '#FFFFFF',
  border: '#000000'
};
const TILE = {w:96, h:96, gap:8};
const CARD = {h:140, padX:24, padY:18, gap:16};
const FONT = 'Nunito Sans';

// --- Menu -----------------------------------------------------------------
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WRS Slides')
    .addItem('Show Sidebar','showSidebar')
    .addItem('Build Selected Rows','buildSelected')
    .addItem('Build All (by Substep)','buildAllBySubstep')
    .addItem('Validate Markup (Selected)','validateSelected')
    .addSeparator()
    .addItem('Run Self-Test','wrsSelfTest')
    .addToUi();
}

function showSidebar(){
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('WRS Helper');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Sheet Helpers --------------------------------------------------------
function getSelectedRows(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  return values.map(row => {
    const obj={};
    headers.forEach((h,i)=> obj[h]=row[i]);
    return obj;
  });
}

// --- Build Functions ------------------------------------------------------
function buildSelected(){
  const rows = getSelectedRows();
  if(!rows.length) return;
  buildSlides(rows);
}

function buildAllBySubstep(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const rows = data.map(r=>{
    const obj={};
    headers.forEach((h,i)=> obj[h]=r[i]);
    return obj;
  });
  buildSlides(rows);
}

function buildSlides(rows){
  const deck = SlidesApp.create('WRS Deck');
  rows.forEach(item => {
    const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    const header = `${item.Word} \u00b7 ${item.Substep||''} \u00b7 ${(item.RenderMode||'CARDS')}`;
    const box = safeInsertRect(slide,{x:20,y:20,w:480,h:40});
    if(box){
      box.getText().setText(header);
      styleText(box.getText(),12);
      safeSetBorder(box,1);
    }
    const parsed = parseWord(item.Word||'', item.Markup||'');
    if((item.RenderMode||'CARDS').toUpperCase()==='TILES'){
      renderTiles(slide, parsed.graphemes);
    }else{
      renderCards(slide, parsed, item);
    }
  });
}

// --- Validation -----------------------------------------------------------
function validateSelected(){
  const rows = getSelectedRows();
  const errors=[];
  rows.forEach((r,i)=>{
    try{ parseWord(r.Word||'', r.Markup||''); }
    catch(e){ errors.push(`Row ${i+1}: ${e.message}`); }
  });
  SpreadsheetApp.getUi().alert(errors.length?errors.join('\n'):'No markup errors.');
}

// --- Parsing --------------------------------------------------------------
function parseWord(word, markup){
  const out={graphemes:[], prefix:'', suffix:'', accent:1};
  const pre=/\{prefix=([^}]+)\}/.exec(markup); if(pre) out.prefix=pre[1];
  const suf=/\{suffix=([^}]+)\}/.exec(markup); if(suf) out.suffix=suf[1];
  const acc=/\{accent=(\d)\}/.exec(markup); if(acc) out.accent=parseInt(acc[1],10);
  word=word.replace(/\{macron\}/g,'\u0304')
           .replace(/\{breve\}/g,'\u0306')
           .replace(/\{schwa\}/g,'\u0259');
  const tokens=[];
  let buf='', lock=false;
  for(let i=0;i<word.length;i++){
    const ch=word[i];
    if(ch==='['){lock=true;buf='';continue;}
    if(ch===']'){tokens.push(buf);lock=false;buf='';continue;}
    if(lock){buf+=ch;continue;}
    if(ch==='-'||ch==='.'){ tokens.push(ch); continue; }
    tokens.push(ch);
  }
  out.graphemes=tokens;
  out.word=tokens.filter(t=>t!=='-'&&t!=='.').join('');
  return out;
}

// --- Rendering ------------------------------------------------------------
function renderTiles(slide, graphemes){
  let x=20, y=80;
  graphemes.forEach(g=>{
    if(g==='-'||g==='.') return; // ignore in tiles
    const rect=safeInsertRect(slide,{x,y,w:TILE.w,h:TILE.h});
    if(!rect) return;
    const color = getTileColor(g);
    safeFill(rect,color);
    safeSetBorder(rect,1);
    const text=rect.getText();
    text.setText(g);
    styleText(text,64);
    x+=TILE.w+TILE.gap;
  });
}

function getTileColor(g){
  if(WELDED_SOUNDS.includes(g.toLowerCase())) return COLORS.welded;
  return isVowel(g) ? COLORS.vowel : COLORS.consonant;
}

function renderCards(slide, parsed, item){
  let x=20, y=80;
  if(parsed.prefix){
    const w=measureText(parsed.prefix);
    const rect=safeInsertRect(slide,{x,y,w:hCardWidth(parsed.prefix),h:CARD.h});
    if(rect){
      safeFill(rect,COLORS.affix); safeSetBorder(rect,1);
      rect.getText().setText(parsed.prefix);
      styleText(rect.getText(),64);
      x+=rect.getWidth()+CARD.gap;
    }
  }
  const sylls = parsed.word.split('-');
  sylls.forEach(syl=>{
    const clean=syl.replace(/\./g,'');
    const rect=safeInsertRect(slide,{x,y,w:hCardWidth(clean),h:CARD.h});
    if(!rect) return;
    safeFill(rect,COLORS.card); safeSetBorder(rect,1);
    rect.getText().setText(clean);
    styleText(rect.getText(),64);
    // scoops
    const groups=syl.split('.');
    let scoX=rect.getLeft()+CARD.padX, curX=scoX;
    const charW=(rect.getWidth()-CARD.padX*2)/clean.length;
    groups.forEach(g=>{
      const gW=g.length*charW;
      drawScoop(slide,curX,curX+gW,rect.getTop()+rect.getHeight()-CARD.padY);
      curX+=gW;
    });
    x+=rect.getWidth()+CARD.gap;
  });
  if(parsed.suffix){
    const rect=safeInsertRect(slide,{x,y,w:hCardWidth(parsed.suffix),h:CARD.h});
    if(rect){
      safeFill(rect,COLORS.affix); safeSetBorder(rect,1);
      rect.getText().setText(parsed.suffix);
      styleText(rect.getText(),64);
    }
  }
}

function hCardWidth(text){
  return Math.max(text.length*40 + CARD.padX*2, TILE.w);
}

// --- Drawing --------------------------------------------------------------
function drawScoop(slide,x1,x2,y){
  safeInsertLine(slide,{x1,y1:y,x2,y2:y});
}

// --- Safe Wrappers --------------------------------------------------------
function safeInsertRect(slide,opts){
  try{
    return slide.insertShape(SlidesApp.ShapeType.RECTANGLE,opts.x,opts.y,opts.w,opts.h);
  }catch(e){ return null; }
}
function safeInsertLine(slide,opts){
  try{
    return slide.insertLine(SlidesApp.LineCategory.STRAIGHT,opts.x1,opts.y1,opts.x2,opts.y2);
  }catch(e){ return null; }
}
function safeSetBorder(shape,width){
  try{ shape.getBorder().setWeight(width).getLineFill().setSolidFill(COLORS.border); }catch(e){}
}
function safeFill(shape,color){
  try{ shape.getFill().setSolidFill(color); }catch(e){}
}
function styleText(text,size){
  try{
    text.setFontFamily(FONT); text.setFontSize(size); text.setBold(false); text.setForegroundColor(COLORS.border);
  }catch(e){}
}
function isVowel(g){
  return /^[aeiouy\u0259\u0304\u0306]$/i.test(g.charAt(0));
}
function measureText(t){ return t.length; }

// --- Self Test ------------------------------------------------------------
function wrsSelfTest(){
  const slide = SlidesApp.create('WRS Self-Test').getSlides()[0];
  const rect = safeInsertRect(slide,{x:20,y:20,w:100,h:60});
  if(rect){
    rect.getText().setText('ok');
    styleText(rect.getText(),12);
    safeSetBorder(rect,1);
  }
  drawScoop(slide,20,120,100);
}

// --- Sidebar Interaction --------------------------------------------------
function insertToken(token){
  const sheet=SpreadsheetApp.getActiveSheet();
  const cell=sheet.getActiveCell();
  cell.setValue((cell.getValue()||'')+token);
}
