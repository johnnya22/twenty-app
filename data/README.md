# Editar os dados da Twenty por JSON

O ficheiro lido pela app é `academic-data.json`. A app verifica-o no arranque, quando voltas à janela e através do botão **Admin & dados → Reler**.

`canteen-menu.json` é apenas a cópia de segurança offline da ementa, preço e condições da Cantina da FCT. Em utilização normal, a página **Cantina** consulta diretamente as páginas oficiais da SAS NOVA e atualiza o cache do dispositivo; não mistures este ficheiro com os teus dados académicos.

## Regras essenciais

- Mantém `schemaVersion` em `4`.
- Usa um `id` único e estável em cada item.
- Usa o mesmo `semesterId` em cadeiras e conteúdo do respetivo semestre.
- Usa `courseId` para ligar aulas, materiais, avaliações, perguntas, quizzes e notas à cadeira.
- Usa `scheduleId` para ligar cada aula à ocorrência recorrente correta do horário.
- Usa `componentId` para ligar cada avaliação à componente do método de avaliação.
- Usa `assessmentId` ou `lessonId` para indicar a origem concreta de cada nota.
- Usa `lessonId` ou `lessonIds` para ligar conteúdo a aulas concretas.
- Usa `pastExamId` para ligar uma pergunta ao teste anterior de origem.
- Mantém `meta.syncMode` como `merge` para juntar/atualizar por ID. `replace` substitui todos os metadados locais pelo JSON.
- Depois de preencher o template, podes mudar `meta.isTemplate` para `false`. Mesmo que não mudes, a app reconhece que deixou de estar vazio.

## Horário

`weekday` usa `0` para domingo, `1` para segunda e assim até `6` para sábado. As horas usam `HH:MM`.

```json
{
  "id": "schedule_fisica_t_1",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "weekday": 2,
  "start": "09:00",
  "end": "10:30",
  "type": "T",
  "room": "B1.03"
}
```

Uma aula preparada tem de usar um bloco da mesma cadeira, no mesmo dia da semana e com o mesmo tipo:

```json
{
  "id": "lesson_fisica_04",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "scheduleId": "schedule_fisica_t_1",
  "title": "T04 · Movimento circular",
  "date": "2026-10-06",
  "start": "09:00",
  "end": "10:30",
  "type": "T",
  "room": "B1.03"
}
```

O nome da aula aparece na aula em direto e no modo **Calendário**. `settings.plannerView` aceita `"schedule"`, `"calendar"` ou `"study-day"`; `settings.calendarView` aceita `"day"`, `"three"`, `"week"` ou `"month"`.

## Testes anteriores, perguntas e imagens

Um teste anterior é guardado em `pastExams`; as perguntas apontam para ele através de `pastExamId`:

```json
{
  "id": "past_fisica_2025_t1",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "title": "Teste 1 — época normal",
  "academicYear": "2024/2025",
  "date": "2025-01-15",
  "source": "PDF disponibilizado pelo professor"
}
```

Cada pergunta pode ter imagens separadas para `question`, `solution` e `explanation`. Um caminho relativo torna o conteúdo portátil dentro do projeto; um upload feito na app usa `source: "indexeddb"` e `blobId`.

```json
{
  "id": "past_fisica_2025_t1_q1",
  "pastExamId": "past_fisica_2025_t1",
  "courseId": "course_fisica",
  "lessonIds": ["lesson_fisica_04"],
  "number": "1.1",
  "prompt": "Enunciado transcrito sem alterações",
  "answer": "",
  "explanation": "",
  "images": [
    { "id": "img_q1", "role": "question", "source": "path", "url": "assets/perguntas/q1.png" },
    { "id": "img_s1", "role": "solution", "source": "path", "url": "assets/solucoes/q1.png" }
  ]
}
```

Eventos usam o mesmo formato, com `role: "event"`. A interface de importação aceita ainda o formato abreviado `images: { "question": ["caminho.png"], "solution": [], "explanation": [] }`.

## Planeamento do estudo

Os blocos do modo **Dia de estudo** ficam em `studyBlocks`:

```json
{
  "id": "study_sat_1",
  "semesterId": "sem_2026_27_1",
  "date": "2026-10-10",
  "title": "Rever Movimento circular",
  "start": "09:00",
  "end": "09:50",
  "kind": "study",
  "courseId": "course_fisica",
  "sourceType": "lesson",
  "sourceId": "lesson_fisica_04",
  "completed": false
}
```

`kind` aceita `study`, `break` ou `lunch`. `sourceType` pode ser `task`, `lesson`, `quiz`, `assessment`, `routine` ou `custom`. As revisões semanais ficam em `weeklyReviews`, com `weekStart`, `priorities`, `doubts`, `notes` e `completedAt`.

## Método de avaliação

Os pesos são percentagens. Tipos úteis: `test`, `project`, `exam`, `presentation`, `class` e `other`.

```json
"evaluation": {
  "examReplacesTests": true,
  "replacementPolicy": "if-higher",
  "components": [
    { "id": "fis_tests", "label": "Testes", "count": 2, "weight": 70, "kind": "test", "minimum": 9.5, "replaceable": true },
    { "id": "fis_labs", "label": "Laboratórios", "count": 4, "weight": 30, "kind": "class", "minimum": 10, "replaceable": false },
    { "id": "fis_exam", "label": "Exame", "count": 1, "weight": 0, "kind": "exam", "replaceable": false }
  ]
}
```

`count` indica quantas avaliações pertencem ao grupo; `minimum` é a média mínima desse grupo para aprovação. Para regras de defesa podes usar `defenseEnabled`, `defenseType` (`oral`, `practical` ou `oral-practical`), `defenseThreshold` e `maxWithoutDefense`.

Os campos antigos `examReplacesTests` e `replacementPolicy` continuam suportados. Para uma substituição precisa, usa `replacementAssessmentIds` na avaliação substituta, como no exemplo abaixo.

## Avaliações e notas

Uma avaliação aponta para a componente correta do método:

```json
{
  "id": "fis_teste_1",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "componentId": "fis_tests",
  "type": "Teste",
  "title": "Teste 1",
  "date": "2026-11-03",
  "time": "10:00",
  "lessonIds": ["lesson_fisica_04"],
  "requiresTestSheet": true,
  "openBook": false,
  "hasDefense": false,
  "replacementAssessmentIds": [],
  "replacementPolicy": "if-higher"
}
```

Um exame pode selecionar exatamente as avaliações que substitui:

```json
{
  "id": "fis_exame",
  "courseId": "course_fisica",
  "componentId": "fis_exam",
  "type": "Exame",
  "title": "Exame",
  "replacementAssessmentIds": ["fis_teste_1", "fis_teste_2"],
  "replacementPolicy": "if-higher"
}
```

A nota fica ligada à avaliação — não apenas à cadeira:

```json
{
  "id": "grade_fis_teste_1",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "assessmentId": "fis_teste_1",
  "componentId": "fis_tests",
  "score": 18.5,
  "date": "2026-11-10",
  "defenseStatus": "completed",
  "defenseType": "oral",
  "defenseFinalScore": 18
}
```

Para uma nota dada numa aula, usa `lessonId`. Só entra na média se a cadeira tiver uma componente `kind: "class"`; caso contrário fica como registo associado à aula.

## PDFs

Ficheiros enviados através da app ficam no IndexedDB do navegador e o JSON guarda apenas `blobId`; os bytes não fazem parte do backup JSON. Para materiais portáteis no projeto, coloca o PDF em `assets/slides/` e usa `source: "url"` com um caminho relativo.

Consulta `academic-data.example.json` para um conjunto completo, claramente identificado como demonstração.

## BEFIRST™ e quizzes de aula

Uma aula revista pode guardar a data de conclusão do quiz:

```json
{
  "id": "lesson_fisica_04",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "title": "T04 · Movimento circular",
  "date": "2026-10-08",
  "start": "09:00",
  "end": "10:30",
  "type": "T",
  "quizCompletedAt": "2026-10-08T18:22:00.000Z"
}
```

Também fica concluída automaticamente quando um quiz com o mesmo `lessonId` recebe `lastCompletedAt`.

Quizzes gerados a partir do banco de perguntas usam `generatedFromPastQuestions: true`. Cada pergunta reutilizada mantém `sourceQuestionId`, para a app não a duplicar:

```json
{
  "id": "quiz_fisica_04",
  "semesterId": "sem_2026_27_1",
  "courseId": "course_fisica",
  "lessonId": "lesson_fisica_04",
  "title": "Quiz da aula · Movimento circular",
  "generatedFromPastQuestions": true,
  "questions": [
    {
      "id": "quizq_fisica_04_1",
      "sourceQuestionId": "past_fisica_2025_t1_q3",
      "sourceType": "past-test",
      "mode": "self-check",
      "prompt": "Enunciado real da pergunta",
      "answer": "Solução guardada pelo administrador",
      "academicYear": "2025/26",
      "assessmentLabel": "Teste 1 · pergunta 3"
    }
  ]
}
```

`mode: "self-check"` é usado quando a pergunta original não tem opções reais: o aluno revela a solução e avalia se sabia. Para escolha múltipla, usa `mode: "multiple-choice"`, `options` e `answerIndex`.

A app cria tarefas automáticas com `type: "lesson-quiz"`, `lessonId` e `autoGenerated: true`. Podem ser editadas no JSON, mas normalmente não precisas de as inserir manualmente.
