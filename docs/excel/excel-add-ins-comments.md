---
title: Работа с комментариями с помощью Excel API JavaScript
description: Сведения об использовании API для добавления, удаления и редактирования потоков комментариев и комментариев.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 16569bc1d72391dff0ac35a48e45470ff90852f8
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938719"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>Работа с комментариями с помощью Excel API JavaScript

В этой статье описывается, как добавлять, читать, изменять и удалять комментарии в книге с Excel API JavaScript. Дополнительные новости о функции комментариев можно узнать из комментариев и заметок [в](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) Excel статье.

В API Excel JavaScript комментарий включает как один первоначальный комментарий, так и связанное обсуждение. Он привязан к отдельной ячейке. Любой, кто просматривает книгу с достаточными разрешениями, может ответить на комментарий. Объект [Comment](/javascript/api/excel/excel.comment) сохраняет эти ответы как [объекты CommentReply.](/javascript/api/excel/excel.commentreply) Комментарий следует рассматривать как поток, и в качестве отправной точки в потоке должна быть специальная запись.

![Комментарий Excel с двумя ответами, помеченными "Comment.replies[0]" и "Comment.replies[1].](../images/excel-comments.png)

Комментарии в книге отслеживаются `Workbook.comments` свойством. Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой. Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment). Комментарии также доступны на уровне [таблицы.](/javascript/api/excel/excel.worksheet) Примеры в этой статье работают с комментариями на уровне книг, но их можно легко изменить для использования `Worksheet.comments` свойства.

## <a name="add-comments"></a>Добавление примечаний

Используйте метод `CommentCollection.add` для добавления комментариев в книгу. Этот метод занимает до трех параметров:

- `cellAddress`: Ячейка, в которой добавлен комментарий. Это может быть объект строки [или диапазона.](/javascript/api/excel/excel.range) Диапазон должен быть одной ячейкой.
- `content`. Содержимое комментария. Используйте строку для простых текстовых комментариев. Используйте [объект CommentRichContent](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями.](#mentions)
- `contentType`: [В переименовку ContentType](/javascript/api/excel/excel.contenttype) указывается тип контента. Значение по умолчанию — `ContentType.plain`.

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
> Комментарии, добавленные надстройки, приписываются текущему пользователю этой надстройки.

### <a name="add-comment-replies"></a>Добавление ответов на комментарии

Объект `Comment` — это поток комментариев, содержащий нулевой или более ответов. Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply). Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа. Ответы отображаются в порядке их добавления. Они также приписываются текущему пользователю надстройки.

В следующем примере кода добавляется ответ к первому примечанию в книге.

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>Изменение комментариев

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

Чтобы изменить ответ на комментарий, установите `CommentReply.content` его свойство.

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

Чтобы удалить комментарий, используйте `Comment.delete` метод. Удаление комментария также удаляет ответы, связанные с этим комментарием.

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

## <a name="resolve-comment-threads"></a>Устранение потоков комментариев

Поток комментариев имеет настраиваемое значение boolean, чтобы указать, `resolved` разрешен ли он. Значение `true` означает, что поток комментариев разрешен. Значение `false` означает, что поток комментариев является либо новым, либо вновь открыт.

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

Ответы на комментарии имеют свойство `resolved` readonly. Его значение всегда равно значению остальной части потока.

## <a name="comment-metadata"></a>Метаданные комментариев

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

### <a name="comment-reply-metadata"></a>Метаданные ответа комментариев

В ответах на комментарии хранятся те же типы метаданных, что и первоначальный комментарий.

В следующем примере показано, как отобразить электронную почту, имя автора и дату создания последнего ответа на **комментарий в A2**.

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

## <a name="mentions"></a>Упоминания

[Упоминания используются](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) для тегов коллег в комментарии. Это отправляет им уведомления содержимым вашего комментария. Ваша надстройка может создавать эти упоминания от вашего имени.

Комментарии с упоминаниями необходимо создавать с [помощью объектов CommentRichContent.](/javascript/api/excel/excel.commentrichcontent) Вызов `CommentCollection.add` с одним или несколькими упоминаниями и `CommentRichContent` `ContentType.mention` указанием в качестве `contentType` параметра. Строка `content` также должна быть отформатирована, чтобы вставить упоминание в текст. Формат для упоминания: `<at id="{replyIndex}">{mentionName}</at>` .

> [!NOTE]
> В настоящее время только точное имя упоминания можно использовать в качестве текста ссылки на упоминание. Поддержка сокращенных версий имени будет добавлена позже.

В следующем примере показан комментарий с одним упоминанием.

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

## <a name="comment-events"></a>События комментариев

Надстройка может прослушивать добавления, изменения и удаления комментариев. [События комментариев](/javascript/api/excel/excel.commentcollection#event-details) происходят на `CommentCollection` объекте. Чтобы выслушать события комментариев, зарегистрируйте `onAdded` обработник событий `onChanged` или `onDeleted` комментариев. При обнаружении события комментариев используйте обработник этого события для получения данных о добавленных, измененных или удаленных комментариях. Событие также обрабатывает добавления, изменения и `onChanged` удаления комментариев. 

Каждое событие комментариев запускается только один раз, когда одновременно выполняется несколько дополнений, изменений или удалений. Все объекты [CommentAddedEventArgs,](/javascript/api/excel/excel.commentaddedeventargs) [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)и [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) содержат массивы ID-данных комментариев для карты действий событий обратно в коллекции комментариев.

Дополнительные сведения о регистрации обработчиков событий, обработке событий и удалении обработчиков событий см. в статье [Excel API JavaScript.](excel-add-ins-events.md) 

### <a name="comment-addition-events"></a>События добавления комментариев 
Событие запускается, когда в коллекцию комментариев добавляется один или `onAdded` несколько новых комментариев. Это событие не *запускается* при добавлении ответов в поток комментариев (см. в разделе [События](#comment-change-events) изменения комментариев, чтобы узнать о событиях ответа на комментарии).

В следующем примере показано, как зарегистрировать обработник событий, а затем использовать объект для получения `onAdded` `CommentAddedEventArgs` `commentDetails` массива добавленного комментария.

> [!NOTE]
> Этот пример работает только при добавлении одного комментария. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a>События изменения комментариев 
Событие `onChanged` комментария запускается в следующих сценариях.

- Содержимое комментария обновляется.
- Поток комментариев разрешен.
- Поток комментариев возобновляется.
- Ответ добавляется в поток комментариев.
- Ответ обновляется в потоке комментариев.
- Ответ удаляется в потоке комментариев.

В следующем примере показано, как зарегистрировать обработник событий, а затем использовать объект для получения `onChanged` `CommentChangedEventArgs` `commentDetails` массива измененного комментария.

> [!NOTE]
> Этот пример работает только при смене одного комментария. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a>События удаления комментариев
Событие `onDeleted` запускается при удалении комментария из коллекции комментариев. После удаления комментария его метаданные перестают быть доступны. Объект [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) предоставляет ID-данные комментариев, если надстройка управляет отдельными комментариями.

В следующем примере показано, как зарегистрировать обработник событий, а затем использовать объект для получения `onDeleted` `CommentDeletedEventArgs` `commentDetails` массива удаленных комментариев.

> [!NOTE]
> Этот пример работает только при удалении одного комментария. 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
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
- [Вставьте комментарии и заметки в Excel](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
