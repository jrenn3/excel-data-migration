(function(){const o=document.createElement("link").relList;if(o&&o.supports&&o.supports("modulepreload"))return;for(const e of document.querySelectorAll('link[rel="modulepreload"]'))n(e);new MutationObserver(e=>{for(const t of e)if(t.type==="childList")for(const l of t.addedNodes)l.tagName==="LINK"&&l.rel==="modulepreload"&&n(l)}).observe(document,{childList:!0,subtree:!0});function r(e){const t={};return e.integrity&&(t.integrity=e.integrity),e.referrerPolicy&&(t.referrerPolicy=e.referrerPolicy),e.crossOrigin==="use-credentials"?t.credentials="include":e.crossOrigin==="anonymous"?t.credentials="omit":t.credentials="same-origin",t}function n(e){if(e.ep)return;e.ep=!0;const t=r(e);fetch(e.href,t)}})();const d=document.getElementById("uploadButton"),i=document.getElementById("fileInput");d.addEventListener("click",()=>{i.click()});i.addEventListener("change",c=>{const o=c.target.files[0];o&&(console.log(`File selected: ${o.name}`),s(o))});function s(c){const o=new FormData;o.append("file",c),fetch("http://localhost:5000/upload",{method:"POST",body:o}).then(async r=>{if(!r.ok){const n=await r.text();throw new Error(`Server error: ${n}`)}return r.blob()}).then(r=>{const n=window.URL.createObjectURL(r),e=document.createElement("a");e.href=n,e.download="updated_template.xlsm",document.body.appendChild(e),e.click(),e.remove()}).catch(r=>{console.error("Upload failed:",r.message),alert("Upload failed: "+r.message)})}
