---
title: Работайте с комментариями с помощью API JavaScript для Excel
description: Сведения об использовании API для добавления, удаления и редактирования комментариев и потоков комментариев.
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: 275828915730d3438101315ee28bf76aa8b8bf3f
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890572"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Работайте с комментариями с помощью API JavaScript для Excel

В этой статье описывается, как добавлять, читать, изменять и удалять комментарии в книге с помощью API JavaScript для Excel. Дополнительные сведения о функции комментариев можно узнать в статье [INSERT Comments and notess in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .

В API JavaScript для Excel комментарий включает один начальный комментарий и подключенное обсуждение. Он привязан к отдельной ячейке. Любой пользователь, просматривающий книгу с достаточными разрешениями, может ответить на комментарий. Объект [comment](/javascript/api/excel/excel.comment) хранит эти ответы как объекты [комментрепли](/javascript/api/excel/excel.commentreply) . Обратите внимание на то, что комментарий является потоком и что поток должен иметь специальную запись в качестве отправной точки.

![Комментарий Excel с пометкой "Comment" с двумя ответами, помеченными как "Comment. ответы [0]" и "Comment. ответы [1].](../images/excel-comments.png)

Комментарии в книге отслеживаются `Workbook.comments` свойством. Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой. Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment). Комментарии также доступны на уровне [листа](/javascript/api/excel/excel.worksheet) . Примеры, приведенные в этой статье, работают с комментариями на уровне книги, но их можно легко изменить, `Worksheet.comments` чтобы использовать свойство.

## <a name="add-comments"></a>Добавление примечаний

Используйте `CommentCollection.add` метод, чтобы добавить комментарии в книгу. Этот метод занимает до трех параметров:

- `cellAddress`: Ячейка, в которую добавляется комментарий. Это может быть объект String или [Range](/javascript/api/excel/excel.range) . Диапазон должен быть одной ячейкой.
- `content`: Контент комментария. Используйте строку для примечаний в виде обычного текста. Используйте объект [комментричконтент](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями](#mentions-online-only). 
- `contentType`: Перечисление [ContentType](/javascript/api/excel/excel.contenttype) , определяющее тип контента. Значение по умолчанию — `ContentType.plain`.

В следующем примере кода добавляется примечание в ячейку **A2**.

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> Комментарии, добавленные надстройкой, добавляются к текущему пользователю этой надстройки.

### <a name="add-comment-replies"></a>Добавление ответов на комментарии

`Comment` Объект — это поток комментариев, который содержит ноль или больше ответов. Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply). Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа. Ответы отображаются в порядке их добавления. Они также применяют атрибуты к текущему пользователю надстройки.

В следующем примере кода добавляется ответ к первому примечанию в книге.

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>Редактирование комментариев

Чтобы изменить примечание или ответ на примечание, настройте его свойство `Comment.content` или `CommentReply.content`.

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a>Изменение ответов на комментарии

Чтобы изменить ответ на комментарий, задайте его `CommentReply.content` свойство.

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a>Удалять комментарии.

Чтобы удалить комментарий, `Comment.delete` используйте метод. При удалении комментария также удаляются ответы, связанные с этим комментарием.

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>Удаление ответов на комментарии

Чтобы удалить ответ на комментарий, используйте `CommentReply.delete` метод.

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a>Разрешение потоков комментариев ([Предварительная версия](../reference/requirement-sets/excel-preview-apis.md)) 

Поток комментариев имеет настраиваемое логическое значение `resolved`, которое указывает, разрешено ли оно. Значение `true` означает, что поток комментариев разрешается. Значение `false` означает, что поток комментариев является либо новым, либо повторно открытым.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Ответы на комментарии имеют свойство `resolved` ReadOnly. Его значение всегда равно значению остальной части потока.

## <a name="comment-metadata"></a>Метаданные Comment

Каждое примечание содержит метаданные о его создании, например автора и дату создания. Автором примечаний, созданных вашей надстройкой, считается текущий пользователь.

В следующем примере показано, как отобразить электронную почту автора, имя автора и дату создания примечания в ячейке **A2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a>Метаданные ответа на комментарии

Ответы на комментарии хранят те же типы метаданных, что и исходный комментарий.

В приведенном ниже примере показано, как отобразить сообщение об авторе, имя автора и дату создания последнего ответа на комментарий в **ячейке A2**.

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions-online-only"></a>Упоминания ([только в Интернете](../reference/requirement-sets/excel-api-online-requirement-set.md)) 

> [!NOTE]
> API упомянутых комментариев в настоящее время доступны только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> Упоминание комментариев в настоящее время поддерживается только для Excel в Интернете.

[Упоминания](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) используются для обозначения коллег в комментарии. При этом уведомления отправляются с содержимым комментария. Ваша надстройка может создавать эти упоминания от вашего имени.

Комментарии с упоминанием следует создавать с помощью объектов [комментричконтент](/javascript/api/excel/excel.commentrichcontent) . Вызов `CommentCollection.add` с `CommentRichContent` указанием одного или нескольких упоминаний и указанием `ContentType.mention` в `contentType` качестве параметра. `content` Строку также необходимо отформатировать, чтобы вставить упоминание в текст. Формат для упоминания: `<at id="{replyIndex}">{mentionName}</at>`.

> НОТЕ В настоящее время в качестве текста ссылки на упоминание можно использовать только точное имя упоминания. Поддержка сокращенных версий имени будет добавлена позже.

В приведенном ниже примере показан комментарий с одним упоминанием.

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Работа с книгами с использованием API JavaScript для Excel](excel-add-ins-workbooks.md)
- [Вставка комментариев и заметок в Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
