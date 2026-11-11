import {expect,it} from 'vitest';
import {PageFillQueue} from '../src/page/page-fill-queue';

it('prepares at most three pages independently and shares a confirmed commit per ready group',async()=>{
  let active=0,peak=0,revision=4;const stored:number[][]=[];
  const queue=new PageFillQueue<number,number>(async pages=>{
    expect(active).toBe(0);stored.push(pages);
    await new Promise(resolve=>setTimeout(resolve,1));return pages.map(()=>++revision);
  });
  const jobs=Array.from({length:17},(_,i)=>queue.run(i,async()=>{
    active++;peak=Math.max(peak,active);await new Promise(resolve=>setTimeout(resolve,1));active--;return i;
  }));
  expect((await Promise.all(jobs)).length).toBe(17);
  expect(peak).toBe(3);expect(stored.every(p=>p.length<=3)).toBe(true);
  expect(stored.flat().sort((a,b)=>a-b)).toEqual(Array.from({length:17},(_,i)=>i));
  expect(queue.pending).toBe(false);
});

it('does not claim a saved result before commit succeeds; a failed commit rejects every affected page',async()=>{
  let fail:(error:Error)=>void;const blocker=new Promise<number[]>((_,reject)=>{fail=reject;});
  const queue=new PageFillQueue<number,number>(()=>blocker);
  let returned=0;
  const jobs=[1,2,3].map(key=>queue.run(key,async()=>key).then(()=>{returned++;},e=>e.message));
  await new Promise(resolve=>setTimeout(resolve,0));expect(returned).toBe(0);
  fail!(new Error('storage unavailable'));expect(await Promise.all(jobs)).toEqual(Array(3).fill('storage unavailable'));
  expect(returned).toBe(0);
});

it('rejects same-page overlap and keeps successful peers after a preparation failure',async()=>{
  let finish:(value:number)=>void;const slow=new Promise<number>(resolve=>{finish=resolve;});
  const committed:number[][]=[];
  const queue=new PageFillQueue<number,number>(async values=>{committed.push(values);return values;});
  const first=queue.run(1,()=>slow);
  await expect(queue.run(1,async()=>1)).rejects.toThrow('PAGE_BUSY');
  const failed=queue.run(2,async()=>{throw new Error('Locked element changed: brand');});
  const assertion=expect(failed).rejects.toThrow('Locked element changed');
  const third=queue.run(3,async()=>3);
  finish!(1);await assertion;expect(await Promise.all([first,third])).toEqual([1,3]);
  expect(committed).toEqual([[1,3]]);
});
