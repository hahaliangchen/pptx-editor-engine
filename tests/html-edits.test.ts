import {expect,it} from 'vitest';
import {editHtmlSource} from '../src/design/html-edits';

it('changes only the exact source fragment and supports insertion and deletion',()=>{
  const source='<body><img data-locked="all" src="fixed"><p>Old title</p><div>remove</div></body>';
  const result=editHtmlSource(source,[{old_string:'Old title',new_string:'New title'},
    {old_string:'<div>remove</div>',new_string:''},{old_string:'</body>',new_string:'<p>Added</p></body>'}]);
  expect(result).toBe('<body><img data-locked="all" src="fixed"><p>New title</p><p>Added</p></body>');
});
it('rejects stale or ambiguous anchors instead of guessing a replacement',()=>{
  expect(()=>editHtmlSource('same same',[{old_string:'same',new_string:'x'}])).toThrow('HTML_EDIT_AMBIGUOUS');
  expect(()=>editHtmlSource('<p>latest manual edit</p>',[{old_string:'old text',new_string:'x'}])).toThrow('HTML_EDIT_NOT_FOUND');
  expect(()=>editHtmlSource('page',[{old_string:'',new_string:'x'}])).toThrow('HTML_EDIT_INVALID');
});
it('shows where a mistyped anchor leaves the real source',()=>{
  // 2026-11-12: the model escaped an attribute's closing quote as &quot;.
  const source='<span style="font-family:&quot;Noto Sans CJK SC&quot;;letter-spacing:0px">03</span>';
  expect(()=>editHtmlSource(source,[{old_string:'letter-spacing:0px&quot;>03</span>',new_string:'x'}]))
    .toThrow('then reads: "letter-spacing:0px\\">03</span>"');
});
