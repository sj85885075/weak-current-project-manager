/* Retained drawing layers: the CAD image survives viewport and device edits. */
(() => {
  const NS='http://www.w3.org/2000/svg';
  class PointCanvasRenderer {
    constructor(svg){this.svg=svg;this.frame=0;this.layers={};}
    initialize(defs){if(this.layers.background)return;this.svg.innerHTML=defs;for(const name of ['background','lines','regions','devices','controls','draft']){const g=document.createElementNS(NS,'g');g.dataset.pmLayer=name;if(name==='background'||name==='regions'){g.setAttribute('pointer-events','none');g.setAttribute('aria-hidden','true');}this.svg.appendChild(g);this.layers[name]=g;}}
    layer(name,markup){const g=this.layers[name];if(g._markup===markup)return;g.innerHTML=markup;g._markup=markup;}
    schedule(fn){this.next=fn;if(this.frame)return;this.frame=requestAnimationFrame(()=>{this.frame=0;const task=this.next;this.next=null;if(this.svg.clientWidth)task?.();});}
    preview(device,lines){
      const node=Array.from(this.layers.devices.children).find(n=>n.dataset.id===device.id);if(!node)return;
      const x=Number(node.dataset.x),y=Number(node.dataset.y);node.setAttribute('transform',`translate(${device.x-x} ${y-device.y})`);
      this.layers.devices._markup=null;
      node.querySelector('use')?.setAttribute('transform',`translate(${x} ${-y}) rotate(${device.orientation||0}) scale(${device.flip_x?-1:1} ${device.flip_y?-1:1}) translate(${-x} ${y})`);
      for(const line of lines){if(line.from_id!==device.id&&line.to_id!==device.id)continue;const g=Array.from(this.layers.lines.children).find(n=>n.dataset.id===line.id);if(!g)continue;this.layers.lines._markup=null;const points=line.points.map(p=>p[0]+','+(-p[1])).join(' ');g.querySelectorAll('polyline').forEach(p=>p.setAttribute('points',points));}
    }
  }
  // Use drawing coordinates, independent of zoom and visible category filters.
  window.connectNearestEndpoints=function(points,devices,fromId='',toId=''){
    const valid=devices.filter(d=>Number.isFinite(d.x)&&Number.isFinite(d.y));
    const byId=new Map(valid.map(d=>[d.id,d]));
    const nearest=(pt,exclude)=>valid.reduce((best,d)=>d.id===exclude?best:!best||Math.hypot(d.x-pt[0],d.y-pt[1])<Math.hypot(best.x-pt[0],best.y-pt[1])?d:best,null);
    const result={points:points.map(p=>[...p]),from_id:fromId,to_id:toId};
    if(points.length<2)return result;
    let from=byId.get(fromId),to=byId.get(toId);
    if(!from)from=nearest(points[0],to?.id);
    if(!to)to=nearest(points.at(-1),from?.id);
    if(from){result.from_id=from.id;result.points[0]=[from.x,from.y];}
    if(to){result.to_id=to.id;result.points[result.points.length-1]=[to.x,to.y];}
    return result;
  };
  window.PointCanvasRenderer=PointCanvasRenderer;
})();
