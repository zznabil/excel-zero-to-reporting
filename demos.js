/* Fixed-rule teaching demos, not an arbitrary Excel formula engine. No eval. */
window.EXCEL_DEMOS = (() => {
  const address = text => {
    const m=/^([A-Z]+)([1-9]\d*)$/.exec(text);
    if(!m)throw new Error('Invalid example cell address.');
    return {column:[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),row:Number(m[2])};
  };
  const name=(column,row)=>{let s='';while(column){column--;s=String.fromCharCode(65+column%26)+s;column=Math.floor(column/26);}return s+row;};
  const grid=(cells,range)=>{const [a,b=a]=range.split(':');const start=address(a),end=address(b);return Array.from({length:end.row-start.row+1},(_,r)=>Array.from({length:end.column-start.column+1},(_,c)=>cells[name(start.column+c,start.row+r)]??''));};
  const values=(cells,range)=>grid(cells,range).flat();
  const eq=(a,b)=>typeof a==='string'&&typeof b==='string'?a.toLowerCase()===b.toLowerCase():a===b;
  const criteriaEqual=(value,criterion)=>{
    if(typeof criterion==='string'&&(/[*?~]/.test(criterion)||/^[<>=]/.test(criterion)))throw new Error('This guided demo supports literal criteria only. Wildcards and comparison text must be tried in real Excel.');
    return eq(value,criterion);
  };
  const number=x=>{if(typeof x!=='number'||!Number.isFinite(x))throw new Error('This calculation needs numeric inputs.');return x;};
  const numericText=value=>{const text=String(value).trim();if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)||!Number.isFinite(Number(text)))throw new Error('Use plain numeric text, without currency symbols or grouping (#VALUE! in Excel).');return Number(text);};
  const sum=xs=>xs.filter(x=>typeof x==='number').reduce((a,b)=>a+b,0);
  const roundingOps=new Set(['roundProduct','matrixRates','sumRounded','gross']);
  function inputValue(example,ref,text){
    if(typeof example.cells[ref]!=='number')return text;
    const value=Number(text);
    if(text.trim()===''||!Number.isFinite(value)||Math.abs(value)>1e9)throw Error(`Enter a number between -1000000000 and 1000000000 in ${ref}.`);
    return roundingOps.has(example.rule.op)?text:value;
  }
  const roundingLimit=detail=>new Error('Guided ROUND limit: '+detail+' Try this case in real Excel; this browser result is unsupported.');
  function normaliseDecimal(units,power){
    if(units===0n)return {units:0n,power:0};
    while(units%10n===0n){units/=10n;power++;}
    if((units<0n?-units:units).toString().length>15)throw roundingLimit('Each input and exact intermediate must fit 15 significant digits.');
    if(power< -324||power>308)throw roundingLimit('The normalised decimal scale must be between -324 and 308.');
    return {units,power};
  }
  function decimal(value){
    const text=typeof value==='number'?String(number(value)):String(value).trim();
    if(text.length>400)throw roundingLimit('Use at most 400 input characters.');
    numericText(text);
    const [mantissa,exponent='0']=text.toLowerCase().split('e'),power=Number(exponent);
    if(Math.abs(power)>1000)throw roundingLimit('Use an explicit exponent between -1000 and 1000.');
    const point=mantissa.indexOf('.'),places=point<0?0:mantissa.length-point-1;
    return normaliseDecimal(BigInt(mantissa.replace('.','')),power-places);
  }
  const decimalProduct=(a,b)=>normaliseDecimal(a.units*b.units,a.power+b.power);
  const decimalSum=(a,b)=>{
    const power=Math.min(a.power,b.power);
    return normaliseDecimal(a.units*10n**BigInt(a.power-power)+b.units*10n**BigInt(b.power-power),power);
  };
  function round(value,digits){
    const shift=value.power+digits;
    if(shift>=0)return value;
    const divisor=10n**BigInt(-shift),absolute=value.units<0n?-value.units:value.units;
    const rounded=absolute/divisor+(absolute%divisor*2n>=divisor?1n:0n);
    return normaliseDecimal(value.units<0n?-rounded:rounded,-digits);
  }
  const decimalNumber=value=>Number(value.units.toString()+'e'+value.power);
  const serial=date=>(date.getTime()-Date.UTC(1899,11,30))/86400000;
  const dateForDay=(rule,cells)=>{const day=number(cells[rule.endDay]);if(!Number.isInteger(day)||day<1||day>31)throw new Error('Use a whole January day from 1 to 31.');return new Date(Date.UTC(rule.year,rule.month-1,day));};
  function calculate(example,cells){
    const r=example.rule;let result;
    const get=ref=>cells[ref]??'';
    const matches=(criteria,index)=>criteria.every(c=>{const item=values(cells,c.range)[index],key=get(c.cell);return c.compare==='>='?number(item)>=number(key):criteriaEqual(item,key);});
    switch(r.op){
      case 'identity':result=get(r.ref);break;
      case 'sum':result=r.refs?sum(r.refs.map(get)):sum(values(cells,r.range));break;
      case 'product':result=r.refs.reduce((n,ref)=>n*number(get(ref)),1);break;
      case 'ratioSum':if(number(get(r.denominator))===0)throw new Error('The denominator is zero (#DIV/0! in Excel).');result=sum(values(cells,r.range))/get(r.denominator);break;
      case 'sumPlus':result=sum(values(cells,r.range))+number(get(r.add));break;
      case 'counta':result=values(cells,r.range).filter(x=>x!=='').length;break;
      case 'label':result=r.prefix+get(r.ref);break;
      case 'sumifs':case 'averageifs':{const selected=values(cells,r.sum).filter((_,i)=>matches(r.criteria,i));if(r.op==='averageifs'&&!selected.length)throw new Error('No matching numeric records (#DIV/0! in Excel).');result=r.op==='sumifs'?sum(selected):sum(selected)/selected.length;break;}
      case 'countifs':result=values(cells,r.criteria[0].range).filter((_,i)=>matches(r.criteria,i)).length;break;
      case 'countOr':{const a=values(cells,r.first),b=values(cells,r.second);result=a.filter((v,i)=>eq(v,r.equals)||number(b[i])>=r.minimum).length;break;}
      case 'roundProduct':result=decimalNumber(round(r.refs.reduce((value,ref)=>decimalProduct(value,decimal(get(ref))),decimal(1)),r.digits));break;
      case 'matrixRates':{const rates=values(cells,r.rates).map(decimal);result=grid(cells,r.matrix).map(row=>row.map((v,i)=>decimalNumber(round(decimalProduct(decimal(v),rates[i]),r.digits))));break;}
      case 'sumRounded':result=decimalNumber(values(cells,r.range).reduce((value,v)=>decimalSum(value,round(decimal(v),r.digits)),decimal(0)));break;
      case 'workdays':{const end=dateForDay(r,cells);result=0;for(let d=new Date(Date.UTC(r.year,r.month-1,1));d<=end;d.setUTCDate(d.getUTCDate()+1))if(d.getUTCDay()!==0&&d.getUTCDay()!==6)result++;break;}
      case 'monthEnd':{const offset=number(get(r.offset));if(!Number.isInteger(offset)||Math.abs(offset)>1200)throw new Error('Use a whole month offset between -1200 and 1200.');result=serial(new Date(Date.UTC(r.year,r.month+offset,0)));break;}
      case 'overnight':{const time=text=>{const m=/^([01]\d|2[0-3]):([0-5]\d)$/.exec(text);if(!m)throw new Error('Use a 24-hour time such as 22:00 or 07:30.');return Number(m[1])*60+Number(m[2]);};result=((time(get(r.end))-time(get(r.start))+1440)%1440)/60;break;}
      case 'inclusiveDays':result=(dateForDay(r,cells)-new Date(Date.UTC(r.year,r.month-1,1)))/86400000+1;break;
      case 'ifThreshold':result=(r.compare==='>'?number(get(r.ref))>number(get(r.threshold)):number(get(r.ref))>=number(get(r.threshold)))?r.yes:r.no;break;
      case 'eligibility':result=number(get(r.credits))>=60&&number(get(r.gpa))>=3.5?'Eligible':'Not eligible';break;
      case 'bands':{const n=number(get(r.ref));const i=r.thresholds.findIndex(t=>n>=t);result=r.labels[i<0?r.labels.length-1:i];break;}
      case 'lookup':{const keys=values(cells,r.keys),returns=values(cells,r.values),key=get(r.key);let i=r.mode==='last'?keys.findLastIndex(v=>eq(v,key)):keys.findIndex(v=>eq(v,key));if(i<0&&r.mode==='smaller'){let best=-Infinity;keys.forEach((v,j)=>{if(typeof v==='number'&&typeof key==='number'&&v<=key&&v>best){best=v;i=j;}});}result=i<0?r.missing:returns[i];break;}
      case 'filter':{const criterion=values(cells,r.criteria);const rows=grid(cells,r.data).filter((_,i)=>eq(criterion[i],get(r.key)));result=rows.length?rows:[[r.missing]];break;}
      case 'lookupProduct':{const i=values(cells,r.keys).findIndex(v=>eq(v,get(r.key)));if(i<0)throw new Error('No product match (#N/A in Excel).');result=number(get(r.quantity))*number(values(cells,r.values)[i]);break;}
      case 'groupProduct':result=(number(get(r.first))+number(get(r.second)))*number(get(r.multiplier));break;
      case 'extractNumber':case 'extractLookup':{const text=String(get(r.ref)),at=text.indexOf(r.delimiter);if(at<0)throw new Error('The required delimiter is missing (#N/A in Excel).');const n=numericText(text.slice(0,at));if(r.op==='extractNumber')result=n;else{const i=values(cells,r.keys).findIndex(v=>eq(v,n));result=i<0?r.missing:values(cells,r.values)[i];}break;}
      case 'missing':{const actual=values(cells,r.actual);const missing=values(cells,r.expected).filter(v=>!actual.some(a=>eq(a,v)));result=missing.length?missing.map(v=>[v]):[[r.none]];break;}
      case 'compareNumericText':result=numericText(get(r.first))>numericText(get(r.second))?'Win':'Not win';break;
      case 'uniqueSort':{const unique=[];for(const v of values(cells,r.range))if(!unique.some(u=>eq(u,v)))unique.push(v);result=unique.sort((a,b)=>a===''?1:b===''?-1:String(a).localeCompare(String(b),'en')).map(v=>[v===''?0:v]);break;}
      case 'gross':{const hours=decimal(get(r.hours));if(hours.units<0n)throw new Error('This fictional pay example requires non-negative hours. Use real Excel for negative adjustments or deductions.');result=decimalNumber(round(decimalSum(decimalProduct(hours,decimal(get(r.rate))),decimalProduct(decimal(get(r.sales)),decimal(get(r.commission)))),2));break;}
      case 'share':{const list=values(cells,r.range),denominator=list.filter(v=>v!=='').length;if(!denominator)throw new Error('No populated responses (#DIV/0! in Excel).');result=list.filter(v=>criteriaEqual(v,get(r.key))).length/denominator;break;}
      case 'trim':result=String(get(r.ref)).replace(/ +/g,' ').replace(/^ +| +$/g,'');break;
      case 'sumproduct':{const b=values(cells,r.right);result=sum(values(cells,r.left).map((a,i)=>number(a)*number(b[i])));break;}
      case 'blankCount':result=values(cells,r.range).filter(v=>v==='').length;break;
      case 'textPart':result=String(get(r.ref)).slice(0,r.count);break;
      case 'weighted':{const weights=values(cells,r.weights).map(number),denominator=sum(weights);if(denominator<=0||weights.some(w=>w<0))throw new Error('Weights must be non-negative with a positive total.');result=sum(values(cells,r.values).map((v,i)=>number(v)*weights[i]))/denominator;break;}
      case 'sumTimes':result=sum(values(cells,r.range))*number(get(r.rate));break;
      case 'pmt':{const pv=number(get(r.principal)),annual=number(get(r.annual)),n=number(get(r.periods));if(pv<0||annual<0||annual>1||!Number.isInteger(n)||n<1||n>1200)throw new Error('Use a non-negative principal, annual rate from 0 to 1, and 1–1200 whole monthly periods.');const rate=annual/12;result=rate===0?pv/n:pv*rate/(-Math.expm1(-n*Math.log1p(rate)));break;}
      case 'breakEven':result=number(get(r.units))*number(get(r.margin))-number(get(r.fixed));break;
      default:throw new Error('This guided example has an unsupported rule.');
    }
    const matrix=Array.isArray(result)?result:[[result]];
    if(matrix.flat().some(v=>v===undefined||typeof v==='number'&&!Number.isFinite(v)))throw new Error('The example could not produce a valid result.');
    return matrix;
  }
  return {address,name,grid,calculate,inputValue};
})();
