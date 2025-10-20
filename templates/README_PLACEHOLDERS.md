Place your ministry DOCX template as `templates/ministry_template.docx`.

Supported placeholders in the template (docxtemplater syntax):

- {{teacher_name}}
- {{subject}}
- {{grade}}
- {{week}}
- {{lesson_title}}
- {{assessment}}
- {{homework}}
- {{today}}  (تاريخ اليوم - اختياري)

For array fields, use a loop in the DOCX:

- Objectives:
  {#objectives}
  - { . }
  {/objectives}

- Materials:
  {#materials}
  - { . }
  {/materials}

- Steps:
  {#steps}
  - { . }
  {/steps}

Notes:
- The dot `{ . }` prints each array item. You can also use `{this}`.
- Ensure your DOCX is saved in UTF-8 compatible fonts for Arabic script.
- Keep the file name exactly `ministry_template.docx`.
