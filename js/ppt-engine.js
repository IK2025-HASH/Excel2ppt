/* ===== ppt-engine.js =====
   PptxGenJS generation: Detail Plan, QA PoaP (paginated), Exec Dashboard.
   Depends on: sparrow-utils.js, generator-core.js (state, computeProgrammeWindow,
               getMilestones, getBands, getIncludedSelectedRowsForFc, getAllRowsFlatIncluded, isQaRow),
               PptxGenJS (CDN global).
   ===== */

/* ========= PPT context builder ========= */
function pptBuildCtx(pptx){
  const {sDate,eDate,sYear,eYear}=computeProgrammeWindow();
  const totMs=Math.max(1,eDate-sDate);

  const W=13.33,H=7.5;
  const LW=2.25,RW=0.25;
  const GX=LW,GW=W-LW-RW;

  const titleH=0.28,yearH=0.22,monthH=0.20;
  const milesH=getMilestones().length?0.17:0;
  const top=titleH+yearH+monthH+milesH;
  const bottom=0.35;
  const GH=H-top-bottom;

  const d2x=(d)=>{
    if(!(d instanceof Date)||isNaN(d))return GX;
    const ms=Math.max(0,Math.min(d-sDate,totMs));
    return GX+(ms/totMs)*GW;
  };

  return {pptx,ST:pptx.ShapeType,W,H,LW,GX,GW,titleH,yearH,monthH,milesH,top,GH,sDate,eDate,sYear,eYear,d2x,milestones:getMilestones(),bands:getBands()};
}

/* ========= Header (title + year + month + milestones) ========= */
function pptDrawHeader(slide,title,ctx){
  const ST=ctx.ST;
  slide.background={color:'FFFFFF'};

  /* Title bar */
  slide.addShape(ST.rect,{x:0,y:0,w:ctx.W,h:ctx.titleH,fill:{color:'CC0000'},line:{color:'CC0000'}});
  slide.addText(title,{x:0,y:0,w:ctx.W,h:ctx.titleH,fontFace:'Calibri',fontSize:12,bold:true,color:'FFFFFF',align:'center',valign:'mid'});

  /* Year bar */
  const yY=ctx.titleH;
  slide.addShape(ST.rect,{x:0,y:yY,w:ctx.W,h:ctx.yearH,fill:{color:'0D4B6E'},line:{color:'0D4B6E'}});
  for(let yr=ctx.sYear;yr<=ctx.eYear;yr++){
    const x1=Math.max(ctx.d2x(new Date(yr,0,1)),ctx.GX);
    const x2=Math.min(ctx.d2x(new Date(yr+1,0,1)),ctx.GX+ctx.GW);
    const w=x2-x1;
    if(w>0.05){
      slide.addText(String(yr),{x:x1,y:yY,w:w,h:ctx.yearH,fontFace:'Calibri',fontSize:10,bold:true,color:'FFFFFF',align:'center',valign:'mid'});
    }
  }

  /* Month bar */
  const mY=ctx.titleH+ctx.yearH;
  slide.addShape(ST.rect,{x:0,y:mY,w:ctx.W,h:ctx.monthH,fill:{color:'1565C0'},line:{color:'1565C0'}});
  slide.addText('Epic',{x:0,y:mY,w:ctx.LW,h:ctx.monthH,fontFace:'Calibri',fontSize:7,bold:true,color:'FFFFFF',align:'center',valign:'mid'});

  const cur=new Date(ctx.sDate.getFullYear(),ctx.sDate.getMonth(),1);
  while(cur<=ctx.eDate){
    const x=ctx.d2x(cur);
    if(x>=ctx.GX&&x<=ctx.GX+ctx.GW){
      slide.addShape(ST.line,{x:x,y:mY,w:0,h:ctx.monthH,line:{color:'FFFFFF',transparency:45,width:0.75}});
      slide.addText(MONTHS[cur.getMonth()],{x:x+0.03,y:mY,w:0.5,h:ctx.monthH,fontFace:'Calibri',fontSize:6,color:'FFFFFF',align:'left',valign:'mid'});
    }
    cur.setMonth(cur.getMonth()+1);
  }

  /* Milestones */
  if(ctx.milesH>0){
    const msY=ctx.titleH+ctx.yearH+ctx.monthH;
    slide.addShape(ST.rect,{x:ctx.GX,y:msY,w:ctx.GW,h:ctx.milesH,fill:{color:'E8EDF4'},line:{color:'E8EDF4'}});
    ctx.milestones.forEach(m=>{
      const bx=Math.max(ctx.d2x(m.start),ctx.GX);
      const ex=Math.min(ctx.d2x(m.end),ctx.GX+ctx.GW);
      const bw=ex-bx;
      if(bw<=0.05)return;
      const mc=pptHex(m.color);
      slide.addShape(ST.roundRect,{x:bx,y:msY+0.01,w:bw,h:ctx.milesH-0.02,fill:{color:mc},line:{color:mc},radius:0.06});
      slide.addText(m.label,{x:bx,y:msY+0.01,w:bw,h:ctx.milesH-0.02,fontFace:'Calibri',fontSize:6,bold:true,color:contrast(mc),align:'center',valign:'mid'});
    });
  }
}

/* ========= Vertical bands overlay ========= */
function pptDrawVerticalBands(slide,ctx,y,h){
  const ST=ctx.ST;
  ctx.bands.forEach(b=>{
    const bx=Math.max(ctx.d2x(b.start),ctx.GX);
    const ex=Math.min(ctx.d2x(b.end),ctx.GX+ctx.GW);
    const bw=ex-bx;
    if(bw<=0.05)return;
    const bc=pptHex(b.color);
    slide.addShape(ST.rect,{x:bx,y:y,w:bw,h:h,fill:{color:bc,transparency:88},line:{color:bc,transparency:65,width:1}});
  });
}

/* ========= Month grid lines ========= */
function pptDrawMonthGridLines(slide,ctx,y,h){
  const ST=ctx.ST;
  const cur=new Date(ctx.sDate.getFullYear(),ctx.sDate.getMonth(),1);
  while(cur<=ctx.eDate){
    const x=ctx.d2x(cur);
    if(x>=ctx.GX&&x<=ctx.GX+ctx.GW){
      slide.addShape(ST.line,{x:x,y:y,w:0,h:h,line:{color:'D8D8D8',width:0.5}});
    }
    cur.setMonth(cur.getMonth()+1);
  }
}

/* ========= Group rows by epic ========= */
function pptGroupByEpic(rows){
  const groups=new Map();
  rows.forEach(r=>{
    const epic=(r.epic||'Activities').trim()||'Activities';
    if(!groups.has(epic))groups.set(epic,[]);
    groups.get(epic).push(r);
  });
  return groups;
}

/* ========= QA PoaP helpers ========= */
function epicGroupsFromRowsOrdered(rows){
  const map=new Map();
  rows.forEach(r=>{
    const epic=(r.epic||'Activities').trim()||'Activities';
    if(!map.has(epic))map.set(epic,[]);
    map.get(epic).push(r);
  });
  return [...map.entries()].map(([epic,arr])=>({epic,rows:arr}));
}

function paginateEpicGroups(groups,maxRowsPerSlide){
  const pages=[];
  let page=[];
  let count=0;

  groups.forEach(g=>{
    let idx=0;
    while(idx<g.rows.length){
      const remaining=maxRowsPerSlide-count;
      if(remaining<=0){
        pages.push(page);
        page=[];
        count=0;
        continue;
      }
      const take=Math.min(remaining,g.rows.length-idx);
      page.push({epic:g.epic,rows:g.rows.slice(idx,idx+take)});
      count+=take;
      idx+=take;

      if(count>=maxRowsPerSlide){
        pages.push(page);
        page=[];
        count=0;
      }
    }
  });

  if(page.length)pages.push(page);
  return pages;
}

/* ========= Draw a timeline slide from epic groups ========= */
function pptDrawTimelineFromEpicGroups(pptx,title,fc,ctx,epicGroups){
  const ST=ctx.ST;
  const slide=pptx.addSlide();
  pptDrawHeader(slide,title,ctx);

  const totalRows=epicGroups.reduce((s,g)=>s+g.rows.length,0);
  if(!totalRows){
    slide.addText('No rows included for current Level/Domain filters.',{x:0.8,y:2.8,w:11.8,h:0.6,fontFace:'Calibri',fontSize:16,bold:true,color:'666666',align:'center',valign:'mid'});
    return;
  }

  const rowH=Math.max(0.18,Math.min(0.28,ctx.GH/Math.max(1,totalRows)));

  let y=ctx.top;
  const leftW=ctx.LW;
  const cEpic=pptHex(fc.col1Color);

  epicGroups.forEach(g=>{
    const gH=g.rows.length*rowH;

    slide.addShape(ST.rect,{x:0,y:y,w:leftW,h:gH,fill:{color:cEpic},line:{color:cEpic}});
    slide.addText(g.epic,{x:0.06,y:y,w:leftW-0.12,h:gH,fontFace:'Calibri',fontSize:8,bold:true,color:contrast(cEpic),align:'center',valign:'mid',wrap:true});

    slide.addShape(ST.rect,{x:ctx.GX,y:y,w:ctx.GW,h:gH,fill:{color:'F7F7F7'},line:{color:'FFFFFF',transparency:100}});
    pptDrawVerticalBands(slide,ctx,y,gH);
    pptDrawMonthGridLines(slide,ctx,y,gH);

    for(const r of g.rows){
      const pct=normPct(r.progress);
      if(r.start&&r.end&&r.end>=r.start){
        const bx=Math.max(ctx.d2x(r.start),ctx.GX);
        const ex=Math.min(ctx.d2x(r.end),ctx.GX+ctx.GW);
        let bw=ex-bx;if(bw<0.05)bw=0.05;

        const labelH=Math.max(0.06,rowH*0.30);
        const barH=Math.max(0.10,rowH*0.45);
        const barY=y+labelH;

        const bc=pptHex(r.color||getBarColorHex(r.activity));
        slide.addShape(ST.roundRect,{x:bx,y:barY,w:bw,h:barH,fill:{color:bc,transparency:(pct>0&&pct<100)?70:0},line:{color:bc,transparency:100},radius:0.06});
        if(pct>0&&pct<100){
          slide.addShape(ST.roundRect,{x:bx,y:barY,w:bw*(pct/100),h:barH,fill:{color:bc},line:{color:bc,transparency:100},radius:0.06});
        }

        // ✅ label ALWAYS: TASK ID - Activity (preserved)
        const label=`${r.id||''} - ${r.activity||''}`.trim();
        const maxW=(ctx.GX+ctx.GW)-bx;
        const labelW=Math.max(1.6,Math.min(Math.max(bw,2.4),maxW));
        slide.addText(label,{x:bx,y:barY,w:labelW,h:barH,fontFace:'Calibri',fontSize:6,bold:true,color:'222222',align:'left',valign:'mid',wrap:false});
      }

      slide.addShape(ST.line,{x:ctx.GX,y:y+rowH,w:ctx.GW,h:0,line:{color:'CCCCCC',width:0.5}});
      y+=rowH;
    }

    slide.addShape(ST.line,{x:0,y:y,w:ctx.GX+ctx.GW,h:0,line:{color:'9E9E9E',width:1}});
  });

  const ts=genTimestamp();
  slide.addText(`Generated: ${ts.display}`,{x:0.2,y:7.22,w:12.9,h:0.25,fontFace:'Calibri',fontSize:7,color:'666666',align:'right'});
}

/* ========= Detail Plan slide (PAGINATED — one or more per workstream) ========= */
function pptDrawDetailSlide(pptx,fc,ctx){
  const rows=getIncludedSelectedRowsForFc(fc);
  if(!rows.length){
    const slide=pptx.addSlide();
    pptDrawHeader(slide,fc.displayName,ctx);
    slide.addText('No rows included for current Level/Domain filters.',{x:0.8,y:2.8,w:11.8,h:0.6,fontFace:'Calibri',fontSize:16,bold:true,color:'666666',align:'center',valign:'mid'});
    return;
  }

  const MIN_ROW_H=0.20;
  const maxRowsPerSlide=Math.max(10,Math.floor(ctx.GH/MIN_ROW_H));
  const epicGroups=epicGroupsFromRowsOrdered(rows);
  const pages=paginateEpicGroups(epicGroups,maxRowsPerSlide);

  pages.forEach((pageGroups,idx)=>{
    const title=pages.length>1
      ?`${fc.displayName} (${idx+1}/${pages.length})`
      :fc.displayName;
    pptDrawTimelineFromEpicGroups(pptx,title,fc,ctx,pageGroups);
  });
}

/* ========= QA PoaP deck (summary + paginated timelines) ========= */
function pptDrawQaSlidesPaginated(pptx,ctx){
  const qaByWs=state.fileConfigs
    .map(fc=>({fc,rows:getIncludedSelectedRowsForFc(fc).filter(isQaRow)}))
    .filter(x=>x.rows.length>0);

  /* Summary slide */
  {
    const slide=pptx.addSlide();
    pptDrawHeader(slide,'QA PoaP (filtered)',ctx);
    const allQa=getAllRowsFlatIncluded().filter(isQaRow);
    slide.addText(`QA rows included: ${allQa.length}`,{x:0.8,y:2.55,w:11.8,h:0.6,fontFace:'Calibri',fontSize:20,bold:true,color:'333333',align:'center',valign:'mid'});
    slide.addText(`Workstreams with QA rows: ${qaByWs.length}`,{x:0.8,y:3.20,w:11.8,h:0.35,fontFace:'Calibri',fontSize:12,color:'666666',align:'center',valign:'mid'});
    slide.addText(`(Auto-pagination enabled)`,{x:0.8,y:3.60,w:11.8,h:0.35,fontFace:'Calibri',fontSize:11,color:'666666',align:'center',valign:'mid'});
  }

  const MIN_ROW_H=0.20;
  const maxRowsPerSlide=Math.max(10,Math.floor(ctx.GH/MIN_ROW_H));

  qaByWs.forEach(({fc,rows})=>{
    const epicGroups=epicGroupsFromRowsOrdered(rows);
    const pages=paginateEpicGroups(epicGroups,maxRowsPerSlide);

    pages.forEach((pageGroups,idx)=>{
      const title=pages.length>1
        ?`QA PoaP — ${fc.displayName} (${idx+1}/${pages.length})`
        :`QA PoaP — ${fc.displayName}`;
      pptDrawTimelineFromEpicGroups(pptx,title,fc,ctx,pageGroups);
    });
  });
}

/* ========= Exec Dashboard slide ========= */
function generateExecSlide(pptx){
  const slide=pptx.addSlide();
  slide.background={color:'0B1020'};

  const ts=genTimestamp();
  const progName=($('progName')?.value||'').trim()||'Programme';
  slide.addText(progName+' – Executive Dashboard',{x:0.6,y:0.6,w:12.5,h:0.6,fontSize:24,fontFace:'Calibri Light',bold:true,color:'FFFFFF'});
  slide.addText('Generated: '+ts.display,{x:0.6,y:1.2,w:12.5,h:0.3,fontSize:10,fontFace:'Calibri',color:'B0BEC5'});

  const rows=getAllRowsFlatIncluded();
  const total=rows.length;
  const completed=rows.filter(r=>normPct(r.progress)>=100).length;
  const inProg=rows.filter(r=>{const p=normPct(r.progress);return p>0&&p<100;}).length;
  const qa=rows.filter(isQaRow).length;

  const ST=pptx.ShapeType;
  function box(x,y,label,val,color){
    slide.addShape(ST.roundRect,{x,y,w:3.9,h:1.15,fill:{color:'11172B'},line:{color:'263238',width:1},radius:0.12});
    slide.addText(label,{x:x+0.18,y:y+0.12,w:3.5,h:0.3,fontFace:'Calibri',fontSize:10,bold:true,color:'90A4AE'});
    slide.addText(String(val),{x:x+0.18,y:y+0.42,w:3.5,h:0.5,fontFace:'Calibri Light',fontSize:24,bold:true,color:color||'00C8E0'});
  }
  box(0.6,2.0,'Total tasks (included)',total,'00C8E0');
  box(4.7,2.0,'Completed',completed,'4CAF50');
  box(8.8,2.0,'In progress',inProg,'FFA726');
  box(0.6,3.4,'QA/Test detected',qa,'AB47BC');
}

/* ========= PPT button handlers ========= */
async function generateDetailPPT(){
  try{
    if(!state.fileConfigs.length){showGenStatus('Load Excel files first.','err');return;}
    if(typeof PptxGenJS==='undefined'){showGenStatus('PptxGenJS not available.','err');return;}
    if(!state.inclusion.taskLevels.size||!state.inclusion.epics.size){showGenStatus('Select at least 1 Level and 1 Domain.','err');return;}

    setProgress(5,'Detail Plan…');
    const pptx=new PptxGenJS();
    pptx.layout='LAYOUT_WIDE';
    const ctx=pptBuildCtx(pptx);

    const total=state.fileConfigs.length;
    state.fileConfigs.forEach((fc,idx)=>{
      pptDrawDetailSlide(pptx,fc,ctx);
      setProgress(5+Math.round(((idx+1)/Math.max(1,total))*85),`Slide ${idx+1}/${total}`);
    });

    setProgress(95,'Writing file…');
    const ts=genTimestamp();
    const safeProg=String(($('progName').value||'Programme')).replace(/[^\w\s\-–]/g,'').trim().replace(/\s+/g,'_');
    const fname=safeProg+'_DetailPlan_'+ts.file+'.pptx';
    await pptx.writeFile({fileName:fname});
    setProgress(100,'Done');
    showGenStatus('Downloaded: '+fname,'ok');
  }catch(e){
    console.error(e);
    showGenStatus('Detail Plan failed: '+(e?.message||String(e)),'err');
  }
}

async function generateQaPPT(){
  try{
    if(!state.fileConfigs.length){showGenStatus('Load Excel files first.','err');return;}
    if(typeof PptxGenJS==='undefined'){showGenStatus('PptxGenJS not available.','err');return;}
    if(!state.inclusion.taskLevels.size||!state.inclusion.epics.size){showGenStatus('Select at least 1 Level and 1 Domain.','err');return;}

    setProgress(10,'QA PoaP…');
    const pptx=new PptxGenJS();
    pptx.layout='LAYOUT_WIDE';
    const ctx=pptBuildCtx(pptx);

    setProgress(25,'Building QA slides…');
    pptDrawQaSlidesPaginated(pptx,ctx);

    setProgress(95,'Writing file…');
    const ts=genTimestamp();
    const safeProg=String(($('progName').value||'Programme')).replace(/[^\w\s\-–]/g,'').trim().replace(/\s+/g,'_');
    const fname=safeProg+'_QA_Poap_'+ts.file+'.pptx';
    await pptx.writeFile({fileName:fname});
    setProgress(100,'Done');
    showGenStatus('Downloaded: '+fname,'ok');
  }catch(e){
    console.error(e);
    showGenStatus('QA PoaP failed: '+(e?.message||String(e)),'err');
  }
}

async function generateExecPPT(){
  try{
    if(typeof PptxGenJS==='undefined'){showGenStatus('PptxGenJS not available.','err');return;}
    if(!state.inclusion.taskLevels.size||!state.inclusion.epics.size){showGenStatus('Select at least 1 Level and 1 Domain.','err');return;}

    setProgress(10,'Exec dashboard…');
    const pptx=new PptxGenJS();
    pptx.layout='LAYOUT_WIDE';
    generateExecSlide(pptx);

    setProgress(95,'Writing file…');
    const ts=genTimestamp();
    const safeProg=String(($('progName').value||'Programme')).replace(/[^\w\s\-–]/g,'').trim().replace(/\s+/g,'_');
    const fname=safeProg+'_ExecDashboard_'+ts.file+'.pptx';
    await pptx.writeFile({fileName:fname});
    setProgress(100,'Done');
    showGenStatus('Downloaded: '+fname,'ok');
  }catch(e){
    console.error(e);
    showGenStatus('Exec failed: '+(e?.message||String(e)),'err');
  }
}