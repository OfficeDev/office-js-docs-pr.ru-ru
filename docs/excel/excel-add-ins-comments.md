---
title: Работа с комментариями с помощью API JavaScript для Excel
description: Сведения об использовании API для добавления, удаления и изменения комментариев и потоков комментариев.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5996c1bb55c3d4a358786b15f7c3e46aae6f42aa
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464799"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Работа с комментариями с помощью API JavaScript для Excel

В этой статье описывается добавление, чтение, изменение и удаление примечаний в книге с помощью API JavaScript для Excel. Дополнительные сведения о функции комментариев см. в статье "Вставка [примечаний и заметок" в Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .

В API JavaScript для Excel комментарий включает как одно начальное примечание, так и подключенное цепочку обсуждений. Он привязан к отдельной ячейке. Любой пользователь, просматривая книгу с достаточными разрешениями, может ответить на комментарий. Объект [Comment](/javascript/api/excel/excel.comment) сохраняет эти ответы как [объекты CommentReply](/javascript/api/excel/excel.commentreply) . Комментарий следует рассматривать как поток, и в качестве отправной точки в потоке должна быть специальная запись.

![Комментарий Excel с двумя ответами с меткой "Comment.replies[0]" и "Comment.replies[1]".](../images/excel-comments.png)

Комментарии в книге отслеживаются свойством `Workbook.comments` . Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой. Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment). Комментарии также доступны на [уровне листа](/javascript/api/excel/excel.worksheet) . Примеры в этой статье работают с комментариями на уровне книги, но их можно легко изменить, чтобы использовать `Worksheet.comments` это свойство.

## <a name="add-comments"></a>Добавление примечаний

Используйте этот `CommentCollection.add` метод для добавления комментариев в книгу. Этот метод принимает до трех параметров:

- `cellAddress`: ячейка, в которую добавляется комментарий. Это может быть строка или [объект Range](/javascript/api/excel/excel.range) . Диапазон должен быть одной ячейкой.
- `content`: содержимое комментария. Используйте строку для комментариев в виде обычного текста. Используйте объект [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями](#mentions).
- `contentType`: [перечисление ContentType](/javascript/api/excel/excel.contenttype) , указывав тип содержимого. Значение по умолчанию — `ContentType.plain`.

В следующем примере кода добавляется примечание в ячейку **A2**.

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> Комментарии, добавленные надстройки, приписываются текущему пользователю надстройки.

### <a name="add-comment-replies"></a>Добавление ответов на комментарии

Объект `Comment` — это поток комментариев, содержащий ноль или больше ответов. Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply). Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа. Ответы отображаются в порядке их добавления. Они также приписыются текущему пользователю надстройки.

В следующем примере кода добавляется ответ к первому примечанию в книге.

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## <a name="edit-comments"></a>Изменение примечаний

Чтобы изменить примечание или ответ на примечание, настройте его свойство `Comment.content` или `CommentReply.content`.

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### <a name="edit-comment-replies"></a>Изменение ответов на комментарии

Чтобы изменить ответ на комментарий, задайте его `CommentReply.content` свойство.

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## <a name="delete-comments"></a>Удалять комментарии.

Чтобы удалить комментарий, используйте `Comment.delete` этот метод. При удалении комментария также удаляются ответы, связанные с этим комментарием.

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### <a name="delete-comment-replies"></a>Удаление ответов на комментарии

Чтобы удалить ответ на комментарий, используйте `CommentReply.delete` метод.

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="resolve-comment-threads"></a>Разрешение потоков примечаний

Поток комментариев имеет настраиваемое логическое значение, указывающее, `resolved`разрешено ли оно. Значение значения означает `true` , что поток комментариев разрешен. Значение означает, что `false` поток примечаний является новым или повторно открыт.

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

Ответы на комментарии имеют свойство только для `resolved` чтения. Его значение всегда равно значению остальной части потока.

## <a name="comment-metadata"></a>Метаданные комментария

Каждое примечание содержит метаданные о его создании, например автора и дату создания. Автором примечаний, созданных вашей надстройкой, считается текущий пользователь.

В следующем примере показано, как отобразить электронную почту автора, имя автора и дату создания примечания в ячейке **A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### <a name="comment-reply-metadata"></a>Комментарий к метаданным ответа

В ответах примечаний хранятся те же типы метаданных, что и в исходном комментарии.

В следующем примере показано, как отобразить сообщение электронной почты автора, имя автора и дату создания последнего ответа на **комментарий в A2**.

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## <a name="mentions"></a>Упоминания

[Упоминания используются](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) для пометки коллег в комментарии. При этом им отправляются уведомления с содержимым комментария. Надстройка может создавать эти упоминания от вашего имени.

Примечания с упоминаниями необходимо создавать с помощью [объектов CommentRichContent](/javascript/api/excel/excel.commentrichcontent) . Вызов `CommentCollection.add` с одним `CommentRichContent` или несколькими упоминаниями и указанием `ContentType.mention` в качестве параметра `contentType` . Строка `content` также должна быть отформатирована для вставки упоминания в текст. Формат упоминания: `<at id="{replyIndex}">{mentionName}</at>`.

> [!NOTE]
> В настоящее время в качестве текста ссылки на упоминание можно использовать только точное имя упоминания. Поддержка сокращенных версий имени будет добавлена позже.

В следующем примере показан комментарий с одним упоминанием.

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## <a name="comment-events"></a>События примечания

Надстройка может ожидать добавления, изменения и удаления примечаний. [События комментария](/javascript/api/excel/excel.commentcollection#event-details) происходят в объекте `CommentCollection` . Чтобы прослушать события примечания, зарегистрируйте `onAdded`обработчик событий или `onChanged``onDeleted` комментариев. При обнаружении события комментария используйте этот обработчик событий для получения данных о добавленном, измененном или удаленном примечании. Событие `onChanged` также обрабатывает добавление, изменение и удаление ответов примечаний.

Каждое событие комментария запускается только один раз при одновременном выполнении нескольких дополнений, изменений или удалений. Все [объекты CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs) и [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) содержат массивы идентификаторов комментариев для сопоставления действий события с коллекциями комментариев.

Дополнительные сведения о регистрации обработчиков событий, обработке событий и удалении обработчиков событий см. в статье "Работа с событиями с помощью [API JavaScript для Excel](excel-add-ins-events.md) ".

### <a name="comment-addition-events"></a>События добавления комментариев

Событие `onAdded` активируется при добавлении одного или нескольких новых комментариев в коллекцию комментариев. Это событие не *активируется* при добавлении ответов в поток комментариев (сведения [](#comment-change-events) о событиях ответа примечания см. в разделе "События изменения комментария").

В следующем примере показано, `onAdded` как зарегистрировать обработчик событий `CommentAddedEventArgs` `commentDetails` , а затем использовать объект для получения массива добавленного комментария.

> [!NOTE]
> Этот пример работает только при добавлении одного комментария.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-change-events"></a>События изменения комментариев

Событие `onChanged` комментария активируется в следующих сценариях.

- Содержимое комментария обновляется.
- Поток комментариев разрешается.
- Поток примечаний снова открывается.
- Ответ добавляется в поток примечаний.
- Ответ обновляется в потоке комментариев.
- Ответ удаляется в потоке комментариев.

В следующем примере показано, `onChanged` как зарегистрировать обработчик `CommentChangedEventArgs` `commentDetails` событий, а затем использовать объект для получения массива измененного комментария.

> [!NOTE]
> Этот пример работает только при изменении одного комментария.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-deletion-events"></a>События удаления примечаний

Событие `onDeleted` активируется при удалении комментария из коллекции комментариев. После удаления комментария его метаданные больше не будут доступны. Объект [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) предоставляет идентификаторы комментариев, если надстройка управляет отдельными комментариями.

В следующем примере показано, `onDeleted` `CommentDeletedEventArgs` `commentDetails` как зарегистрировать обработчик событий, а затем использовать объект для получения массива удаленного комментария.

> [!NOTE]
> Этот пример работает только при удалении одного комментария.

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с книгами с использованием API JavaScript для Excel](excel-add-ins-workbooks.md)
- [Работа с событиями при помощи API JavaScript для Excel](excel-add-ins-events.md)
- [Вставка примечаний и заметок в Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
