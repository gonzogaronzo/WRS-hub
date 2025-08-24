const ACTIVE_SUBSTEP_DEFAULT = "5.4";
async function loadJSON(p){ const r=await fetch(p); return r.json(); }
async function initHome(){
  const subSel=document.getElementById('substep'), badge=document.getElementById('substepBadge');
  if(!subSel||!badge) return;
  const data=await loadJSON('data/scrolls.json');
  const subs=[...new Set(data.map(x=>x.Substep))].sort((a,b)=>a.localeCompare(b));
  subs.forEach(s=>{ const o=document.createElement('option'); o.value=s; o.textContent=s; subSel.appendChild(o); });
  subSel.value=subs.includes(ACTIVE_SUBSTEP_DEFAULT)?ACTIVE_SUBSTEP_DEFAULT:subs[0];
  badge.textContent=subSel.value;
  subSel.addEventListener('change', ()=> badge.textContent=subSel.value);
  document.getElementById('summon').addEventListener('click', ()=> summonDeck(subSel.value));
}
async function summonDeck(substep){
  const [scrolls,hfw,stories]=await Promise.all([
    loadJSON('data/scrolls.json'), loadJSON('data/secret_words.json'), loadJSON('data/story_scrolls.json')
  ]);
  const words=scrolls.filter(r=>r.Substep===substep && r.Category==='word').map(r=>r.Item);
  const sentences=scrolls.filter(r=>r.Substep===substep && r.Category==='sentence').map(r=>r.Item);
  const hfwList=hfw.filter(r=>r.Substep===substep).map(r=>r.Word);
  const story=stories.find(r=>r.Substep===substep);

  const pptx=new PptxGenJS();
  const themeBg='0b0f19', accent='ef4444', ink='e5e7eb';
  function darkSlide(title, bullets){
    const s=pptx.addSlide(); s.background={color:themeBg};
    s.addText(title,{x:0.5,y:0.3,w:9,h:1,fontSize:28,bold:true,color:accent,fontFace:'Arial'});
    if(bullets?.length){
      s.addText(bullets.map(t=>'• '+t).join('\n'),{x:0.7,y:1.3,w:8.6,h:4.5,fontSize:18,color:ink,fontFace:'Arial'});
    }
  }
  darkSlide(`🥷 Wilson Dojo — Lesson ${substep}`, [
    'Parts 1–10 auto‑generated from JSON',
    hfwList.length?('HFW: '+hfwList.join(', ')):'HFW: —'
  ]);
  darkSlide('🗡️ Part 1–2  Quick Drill / Teach‑Review', [
    'Letter‑keyword‑sound review (no new phonemes in 5.4).',
    'Model syllable division; note schwa in open syllables.'
  ]);
  darkSlide('🗡️ Part 3  Word Cards', words.slice(0,10));
  darkSlide('🗡️ Part 4  Wordlist Reading', words.slice(0,10));
  darkSlide('🗡️ Part 5  Sentence Reading', sentences.slice(0,6));
  darkSlide('🗡️ Part 6–8  Spelling & Dictation', [
    'Reverse drill of sounds/word‑elements.',
    'Spell multisyllabic words; then write.',
    'Dictation: repeat, circle HFW, proofread.'
  ]);
  darkSlide('🗡️ Part 9  Controlled Text', story?[
    `Reader: ${story.Title} (${story.ReaderID})`,
    story.Summary||'',
    story.Link?('Link: '+story.Link):'(Add link in story_scrolls.json)'
  ]:['(Add a story to story_scrolls.json)']);
  darkSlide('🗡️ Part 10  Comprehension & Vocabulary', [
    'Retell events; phrasing and expression.',
    'Vocabulary from complex base words; discuss meaning.'
  ]);
  darkSlide('Lesson Ready',[
    'This deck pulls from /data JSON.',
    'Edit the JSON → re‑download a fresh deck.'
  ]);
  pptx.writeFile({fileName:`Wilson_Dojo_Lesson_${substep}.pptx`});
}
document.addEventListener('DOMContentLoaded', initHome);
