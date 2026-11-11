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
