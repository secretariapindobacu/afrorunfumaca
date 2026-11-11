// Configure com a URL do Web App publicado no Google Apps Script.
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz5qfS9WS1lTVpmZxwrezPGdsGUe9_tmXypDBzfcN5VczF7HHNQVwFC2Nfwtv_NA2XCpw/exec";
// INSCRIÇÕES: com MODO_TESTE = true o formulário fica sempre aberto (para testar).
// Com MODO_TESTE = false ele fica fechado até a data de abertura e abre sozinho.
// Mesmo fechado, dá para testar abrindo o site com ?teste=1 no final do link.
const MODO_TESTE = true;
const ABERTURA_INSCRICOES = new Date("2026-11-12T00:00:00-03:00");
const EVENT_DATE = new Date("2026-11-28T15:30:00-03:00");
const form = document.getElementById("registrationForm");
const errorBox = document.getElementById("formError");
const confirmation = document.getElementById("confirmation");
let currentBib = null;

const pad = n => String(n).padStart(2, "0");
function updateCountdown(){
  const diff = Math.max(0, EVENT_DATE.getTime() - Date.now());
  const days = Math.floor(diff / 86400000);
  const hours = Math.floor(diff % 86400000 / 3600000);
  const minutes = Math.floor(diff % 3600000 / 60000);
  const seconds = Math.floor(diff % 60000 / 1000);
  document.getElementById("days").textContent = pad(days);
  document.getElementById("hours").textContent = pad(hours);
  document.getElementById("minutes").textContent = pad(minutes);
  document.getElementById("seconds").textContent = pad(seconds);
}
updateCountdown(); setInterval(updateCountdown,1000);

const observer = new IntersectionObserver(entries => entries.forEach(entry => {
  if(entry.isIntersecting){entry.target.classList.add("visible");observer.unobserve(entry.target);}
}), {threshold:.12});
document.querySelectorAll(".reveal").forEach(el => observer.observe(el));

function updateTrack(){
  const track=document.getElementById("track");
  const rect=track.getBoundingClientRect();
  const vh=window.innerHeight;
  const progress=Math.max(0,Math.min(1,(vh-rect.top)/(vh+rect.height)));
  const percent=progress*100;
  document.getElementById("trackProgress").style.width=percent+"%";
  document.getElementById("runner").style.left=percent+"%";
}
window.addEventListener("scroll",updateTrack,{passive:true});updateTrack();

const cpfInput=form.elements.cpf;
cpfInput.addEventListener("input",e=>{
 let v=e.target.value.replace(/\D/g,"").slice(0,11);
 v=v.replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d{1,2})$/,"$1-$2");
 e.target.value=v;
});
form.elements.whatsapp.addEventListener("input",e=>{
 let v=e.target.value.replace(/\D/g,"").slice(0,11);
 if(v.length>10)v=v.replace(/^(\d{2})(\d{5})(\d{4})$/,"($1) $2-$3");
 else if(v.length>6)v=v.replace(/^(\d{2})(\d{4})(\d{1,4})$/,"($1) $2-$3");
 else if(v.length>2)v=v.replace(/^(\d{2})(\d{1,5})$/,"($1) $2");
 e.target.value=v;
});
function validCPF(value){
 const cpf=value.replace(/\D/g,"");if(cpf.length!==11||/^(\d)\1+$/.test(cpf))return false;
 for(let t=9;t<11;t++){let sum=0;for(let i=0;i<t;i++)sum+=Number(cpf[i])*(t+1-i);let d=(sum*10)%11;if(d===10)d=0;if(d!==Number(cpf[t]))return false;}
 return true;
}
function createBib(name,number){
 return new Promise((resolve,reject)=>{
 const canvas=document.getElementById("bibCanvas"),ctx=canvas.getContext("2d");
 ctx.fillStyle="#FFF4E2";ctx.fillRect(0,0,canvas.width,canvas.height);
 ctx.strokeStyle="#2A140B";ctx.lineWidth=16;ctx.strokeRect(30,30,940,1140);
 const colors=["#FF5A1F","#FFC400","#00B3A7","#F0368A","#2BB24C"];
 colors.forEach((c,i)=>{ctx.fillStyle=c;ctx.fillRect(30+i*188,30,188,20);});
 [[70,70],[930,70],[70,1130],[930,1130]].forEach(([x,y])=>{ctx.beginPath();ctx.arc(x,y,18,0,Math.PI*2);ctx.fillStyle="#bca992";ctx.fill();});
 const img=new Image();img.onload=()=>{
   ctx.drawImage(img,canvas.width/2-190,110,380,420);
   ctx.textAlign="center";ctx.fillStyle="#2A140B";ctx.font="bold 56px Outfit, sans-serif";ctx.fillText("AFRO RUN · FUMAÇA 2026",500,600);
   ctx.fillStyle="#FF5A1F";ctx.font="bold 42px Outfit, sans-serif";ctx.fillText("NÚMERO DE PEITO",500,675);
   ctx.fillStyle="#2A140B";ctx.font="bold 54px Outfit, sans-serif";ctx.fillText(name.toLocaleUpperCase("pt-BR").slice(0,30),500,765);
   ctx.fillStyle="#00B3A7";ctx.font="bold 220px Anton, sans-serif";ctx.fillText(String(number).padStart(4,"0"),500,1010);
   ctx.fillStyle="#2A140B";ctx.font="bold 34px Outfit, sans-serif";ctx.fillText("4 KM  ·  28.11.2026  ·  15H30",500,1090);
   ctx.fillStyle="#F0368A";ctx.font="bold 28px Outfit, sans-serif";ctx.fillText("CULTURA, ESPORTE E IGUALDADE EM MOVIMENTO",500,1140);
   resolve();
 };
 img.onerror=()=>reject(new Error("Não foi possível carregar a logo para o número de peito."));
 img.src="logo-transparente.png";
 });
}
function throwConfetti(){
 const canvas=document.getElementById("confetti"),ctx=canvas.getContext("2d");
 canvas.width=innerWidth*devicePixelRatio;canvas.height=innerHeight*devicePixelRatio;
 canvas.style.width=innerWidth+"px";canvas.style.height=innerHeight+"px";ctx.scale(devicePixelRatio,devicePixelRatio);
 const palette=["#FF5A1F","#FFC400","#00B3A7","#F0368A","#2BB24C"];
 const pieces=Array.from({length:140},()=>({x:Math.random()*innerWidth,y:-Math.random()*innerHeight,size:4+Math.random()*7,speed:2+Math.random()*4,drift:-1.5+Math.random()*3,angle:Math.random()*6,color:palette[Math.floor(Math.random()*palette.length)]}));
 let frame=0;function draw(){ctx.clearRect(0,0,innerWidth,innerHeight);pieces.forEach(p=>{p.y+=p.speed;p.x+=p.drift;p.angle+=.06;ctx.save();ctx.translate(p.x,p.y);ctx.rotate(p.angle);ctx.fillStyle=p.color;ctx.fillRect(-p.size/2,-p.size/2,p.size,p.size*1.5);ctx.restore();});if(++frame<180)requestAnimationFrame(draw);else ctx.clearRect(0,0,innerWidth,innerHeight);}
 if(!matchMedia("(prefers-reduced-motion: reduce)").matches)draw();
}
form.addEventListener("submit",async e=>{
 e.preventDefault();errorBox.textContent="";
 if(!form.reportValidity())return;
 if(!validCPF(form.elements.cpf.value)){errorBox.textContent="Confira o CPF informado.";form.elements.cpf.focus();return;}
 const phone=form.elements.whatsapp.value.replace(/\D/g,"");
 if(phone.length<10){errorBox.textContent="Informe um WhatsApp válido.";form.elements.whatsapp.focus();return;}
 if(!form.elements.nascimento.value){errorBox.textContent="Escolha o dia, o mês e o ano de nascimento.";document.getElementById("nascDia").focus();return;}
 const birth=new Date(form.elements.nascimento.value+"T00:00:00");
 if(isNaN(birth)||birth>=new Date()){errorBox.textContent="Confira a data de nascimento.";return;}
 if(isUnder16()){
   if(form.elements.responsavel.value.trim().length<3){errorBox.textContent="Informe o nome do responsável pelo menor de 16 anos.";form.elements.responsavel.focus();return;}
   if(!form.elements.autorizacao_responsavel.checked){errorBox.textContent="O responsável precisa autorizar e acompanhar o menor de 16 anos (Art. 6º).";return;}
 }
 if(SCRIPT_URL==="COLE_A_URL_DO_WEB_APP_AQUI"){errorBox.textContent="A organização ainda precisa configurar o envio da inscrição. Tente novamente mais tarde.";return;}
 const data=Object.fromEntries(new FormData(form).entries());
 const [ay,am,ad]=data.nascimento.split("-");data.nascimento=ad+"/"+am+"/"+ay;
 data.aptidao=form.elements.aptidao.checked;data.imagem=form.elements.imagem.checked;data.regulamento=form.elements.regulamento.checked;data.menor16=isUnder16();data.autorizacao_responsavel=form.elements.autorizacao_responsavel.checked;if(!data.menor16){data.responsavel="";data.autorizacao_responsavel=false;}
 const submit=form.querySelector("button[type=submit]");submit.disabled=true;submit.textContent="ENVIANDO…";
 try{
   const response=await fetch(SCRIPT_URL,{method:"POST",body:JSON.stringify(data),headers:{"Content-Type":"text/plain;charset=utf-8"}});
   const result=await response.json();
   if(!response.ok||!result.success)throw new Error(result.message||"Não foi possível enviar.");
   currentBib={name:data.nome,number:result.number};
   createBib(currentBib.name,currentBib.number);
   form.hidden=true;confirmation.hidden=false;throwConfetti();confirmation.scrollIntoView({behavior:"smooth",block:"start"});
 }catch(err){errorBox.textContent="Não foi possível concluir. Verifique sua conexão e tente novamente.";submit.disabled=false;submit.textContent="CONFIRMAR INSCRIÇÃO ✳";}
});
document.getElementById("downloadBib").addEventListener("click",()=>{
 const canvas=document.getElementById("bibCanvas");const a=document.createElement("a");a.download="numero-de-peito-afrorun-"+currentBib.number+".png";a.href=canvas.toDataURL("image/png");a.click();
});
document.getElementById("shareBib").addEventListener("click",()=>{
 const message=`Tô dentro da AFRO RUN — Fumaça 2026! 🎉 Meu número de peito é ${currentBib.number}. Cultura, Esporte e Igualdade em Movimento. 4 km, dia 28/11 às 15h30!`;
 window.open("https://wa.me/?text="+encodeURIComponent(message),"_blank","noopener");
});

// ---------- MENU (desktop e celular) ----------
const topbar=document.getElementById("topbar");
const nav=document.getElementById("nav");
const menuToggle=document.getElementById("menuToggle");
function setMenu(open){
 nav.classList.toggle("open",open);
 menuToggle.setAttribute("aria-expanded",String(open));
 menuToggle.setAttribute("aria-label",open?"Fechar menu":"Abrir menu");
 document.body.classList.toggle("menu-open",open);
}
menuToggle.addEventListener("click",()=>setMenu(!nav.classList.contains("open")));
nav.querySelectorAll("a").forEach(a=>a.addEventListener("click",()=>setMenu(false)));
document.addEventListener("keydown",e=>{if(e.key==="Escape"&&nav.classList.contains("open")){setMenu(false);menuToggle.focus();}});
matchMedia("(min-width: 980px)").addEventListener("change",e=>{if(e.matches)setMenu(false);});
window.addEventListener("scroll",()=>topbar.classList.toggle("scrolled",scrollY>10),{passive:true});

// Destaca no menu a seção que está na tela
const navLinks=[...nav.querySelectorAll(":scope > a")];
const sectionObserver=new IntersectionObserver(entries=>entries.forEach(entry=>{
 if(entry.isIntersecting){navLinks.forEach(a=>a.classList.toggle("active",a.getAttribute("href")==="#"+entry.target.id));}
}),{rootMargin:"-45% 0px -50% 0px"});
navLinks.forEach(a=>{const sec=document.querySelector(a.getAttribute("href"));if(sec)sectionObserver.observe(sec);});
sectionObserver.observe(document.getElementById("inicio"));

// ---------- REGULAMENTO ----------
const ruleItems=[...document.querySelectorAll(".rules-list details")];
const toggleRules=document.getElementById("toggleRules");
function syncToggle(){
 const allOpen=ruleItems.every(d=>d.open);
 toggleRules.innerHTML=allOpen?"FECHAR TUDO <b>−</b>":"ABRIR TUDO <b>＋</b>";
}
toggleRules.addEventListener("click",()=>{const open=!ruleItems.every(d=>d.open);ruleItems.forEach(d=>d.open=open);syncToggle();});
ruleItems.forEach(d=>d.addEventListener("toggle",syncToggle));
document.getElementById("printRules").addEventListener("click",()=>{
 const before=ruleItems.map(d=>d.open);
 ruleItems.forEach(d=>d.open=true);
 window.print();
 ruleItems.forEach((d,i)=>d.open=before[i]);syncToggle();
});

// ---------- MENORES DE 16 (Regulamento, Art. 6º) ----------
function isUnder16(){
 const v=form.elements.nascimento.value;if(!v)return false;
 const b=new Date(v+"T00:00:00");const ref=new Date(2026,10,28);
 let age=ref.getFullYear()-b.getFullYear();
 if(ref.getMonth()<b.getMonth()||(ref.getMonth()===b.getMonth()&&ref.getDate()<b.getDate()))age--;
 return age<16;
}
const guardianBox=document.getElementById("guardianBox");
form.elements.nascimento.addEventListener("change",()=>{
 const show=isUnder16();guardianBox.hidden=!show;
 form.elements.responsavel.required=show;form.elements.autorizacao_responsavel.required=show;
});

// ---------- ABERTURA DAS INSCRIÇÕES ----------
(function(){
 const forcarTeste=new URLSearchParams(location.search).has("teste");
 const aberto=MODO_TESTE||forcarTeste||Date.now()>=ABERTURA_INSCRICOES.getTime();
 form.hidden=!aberto;
 document.getElementById("formClosed").hidden=aberto;
})();

// ---------- COMPARTILHAR NO INSTAGRAM (Story / Reels) ----------
const SITE_URL="https://secretariapindobacu.github.io/afrorunfumaca/";
const shareHint=document.getElementById("shareHint");
document.querySelectorAll(".sh-ig").forEach(btn=>btn.addEventListener("click",async()=>{
 const file=btn.dataset.file,type=btn.dataset.type,label=btn.dataset.label;
 shareHint.textContent="Preparando a arte…";btn.setAttribute("aria-busy","true");
 try{
  const blob=await (await fetch(file)).blob();
  const f=new File([blob],file,{type});
  if(navigator.canShare&&navigator.canShare({files:[f]})){
   await navigator.share({files:[f],title:"AFRO RUN Fumaça 2026",text:"Bora pra AFRO RUN Fumaça 2026! Inscrições: "+SITE_URL});
   shareHint.textContent="No menu, escolha Instagram e depois "+label+". Cole o link no adesivo de link: "+SITE_URL;
  }else{
   const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=file;document.body.appendChild(a);a.click();a.remove();
   try{await navigator.clipboard.writeText(SITE_URL);}catch(e){}
   shareHint.textContent="Arte baixada! Abra o Instagram, crie um "+label+" com ela e use o adesivo de link (o link já foi copiado).";
  }
 }catch(err){
  if(err&&err.name==="AbortError"){shareHint.textContent="";}
  else{shareHint.textContent="Não deu para abrir o compartilhamento. Baixe a arte e poste no Instagram: "+location.origin+location.pathname+file;}
 }finally{btn.removeAttribute("aria-busy");}
}));

// ---------- DATA DE NASCIMENTO: dia / mês / ano (anos só para trás) ----------
(function(){
 const dia=document.getElementById("nascDia"),mes=document.getElementById("nascMes"),ano=document.getElementById("nascAno");
 const hidden=form.elements.nascimento;
 const meses=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
 meses.forEach((m,i)=>mes.add(new Option(m,String(i+1).padStart(2,"0"))));
 const atual=new Date().getFullYear();
 for(let y=atual;y>=1920;y--)ano.add(new Option(y,y));
 function diasDoMes(){
  const y=+ano.value||2000,m=+mes.value||1,max=new Date(y,m,0).getDate(),sel=dia.value;
  dia.length=1;for(let d=1;d<=max;d++)dia.add(new Option(d,String(d).padStart(2,"0")));
  if(sel&&+sel<=max)dia.value=sel;
 }
 function sync(){
  hidden.value=(dia.value&&mes.value&&ano.value)?`${ano.value}-${mes.value}-${dia.value}`:"";
  hidden.dispatchEvent(new Event("change"));
 }
 diasDoMes();
 mes.addEventListener("change",()=>{diasDoMes();sync();});
 ano.addEventListener("change",()=>{diasDoMes();sync();});
 dia.addEventListener("change",sync);
})();
