const burger=document.querySelector('.burger'),nav=document.querySelector('.nav nav');
burger.addEventListener('click',()=>{const o=nav.classList.toggle('open');burger.setAttribute('aria-expanded',o)});
nav.addEventListener('click',e=>{if(e.target.tagName==='A'){nav.classList.remove('open');burger.setAttribute('aria-expanded',false)}});
const els=document.querySelectorAll('.section-head,.match article,.card,.timeline li,.stats div,.skills div,.edu li,.award');
els.forEach(e=>e.classList.add('reveal'));
if('IntersectionObserver' in window){
  const io=new IntersectionObserver(es=>es.forEach(x=>{if(x.isIntersecting){x.target.classList.add('in');io.unobserve(x.target)}}),{threshold:.12});
  els.forEach(e=>io.observe(e));
}else els.forEach(e=>e.classList.add('in'));
