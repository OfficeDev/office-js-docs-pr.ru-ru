---
title: Работайте с комментариями с помощью API JavaScript для Excel
description: Сведения об использовании API для добавления, удаления и редактирования комментариев и потоков комментариев.
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: f0be13cc666ed4b6b5b3cfac59f299c872139f4c
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408574"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="8d649-103">Работайте с комментариями с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="8d649-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="8d649-104">В этой статье описывается, как добавлять, читать, изменять и удалять комментарии в книге с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="8d649-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="8d649-105">Дополнительные сведения о функции комментариев можно узнать в статье [INSERT Comments and notess in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="8d649-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="8d649-106">В API JavaScript для Excel комментарий включает один начальный комментарий и подключенное обсуждение.</span><span class="sxs-lookup"><span data-stu-id="8d649-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="8d649-107">Он привязан к отдельной ячейке.</span><span class="sxs-lookup"><span data-stu-id="8d649-107">It is tied to an individual cell.</span></span> <span data-ttu-id="8d649-108">Любой пользователь, просматривающий книгу с достаточными разрешениями, может ответить на комментарий.</span><span class="sxs-lookup"><span data-stu-id="8d649-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="8d649-109">Объект [comment](/javascript/api/excel/excel.comment) хранит эти ответы как объекты [комментрепли](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="8d649-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="8d649-110">Обратите внимание на то, что комментарий является потоком и что поток должен иметь специальную запись в качестве отправной точки.</span><span class="sxs-lookup"><span data-stu-id="8d649-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Комментарий Excel с пометкой "Comment" с двумя ответами, помеченными как "Comment. ответы [0]" и "Comment. ответы [1].](../images/excel-comments.png)

<span data-ttu-id="8d649-112">Комментарии в книге отслеживаются `Workbook.comments` свойством.</span><span class="sxs-lookup"><span data-stu-id="8d649-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="8d649-113">Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="8d649-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="8d649-114">Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="8d649-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="8d649-115">Комментарии также доступны на уровне [листа](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="8d649-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="8d649-116">Примеры, приведенные в этой статье, работают с комментариями на уровне книги, но их можно легко изменить, чтобы использовать `Worksheet.comments` свойство.</span><span class="sxs-lookup"><span data-stu-id="8d649-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="8d649-117">Добавление примечаний</span><span class="sxs-lookup"><span data-stu-id="8d649-117">Add comments</span></span>

<span data-ttu-id="8d649-118">Используйте `CommentCollection.add` метод, чтобы добавить комментарии в книгу.</span><span class="sxs-lookup"><span data-stu-id="8d649-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="8d649-119">Этот метод занимает до трех параметров:</span><span class="sxs-lookup"><span data-stu-id="8d649-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="8d649-120">`cellAddress`: Ячейка, в которую добавляется комментарий.</span><span class="sxs-lookup"><span data-stu-id="8d649-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="8d649-121">Это может быть объект String или [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="8d649-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="8d649-122">Диапазон должен быть одной ячейкой.</span><span class="sxs-lookup"><span data-stu-id="8d649-122">The range must be a single cell.</span></span>
- <span data-ttu-id="8d649-123">`content`: Контент комментария.</span><span class="sxs-lookup"><span data-stu-id="8d649-123">`content`: The comment's content.</span></span> <span data-ttu-id="8d649-124">Используйте строку для примечаний в виде обычного текста.</span><span class="sxs-lookup"><span data-stu-id="8d649-124">Use a string for plain text comments.</span></span> <span data-ttu-id="8d649-125">Используйте объект [комментричконтент](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями](#mentions).</span><span class="sxs-lookup"><span data-stu-id="8d649-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="8d649-126">`contentType`: Перечисление [ContentType](/javascript/api/excel/excel.contenttype) , определяющее тип контента.</span><span class="sxs-lookup"><span data-stu-id="8d649-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="8d649-127">Значение по умолчанию — `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="8d649-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="8d649-128">В следующем примере кода добавляется примечание в ячейку **A2**.</span><span class="sxs-lookup"><span data-stu-id="8d649-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="8d649-129">Комментарии, добавленные надстройкой, добавляются к текущему пользователю этой надстройки.</span><span class="sxs-lookup"><span data-stu-id="8d649-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="8d649-130">Добавление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="8d649-130">Add comment replies</span></span>

<span data-ttu-id="8d649-131">`Comment`Объект — это поток комментариев, который содержит ноль или больше ответов.</span><span class="sxs-lookup"><span data-stu-id="8d649-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="8d649-132">Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="8d649-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="8d649-133">Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа.</span><span class="sxs-lookup"><span data-stu-id="8d649-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="8d649-134">Ответы отображаются в порядке их добавления.</span><span class="sxs-lookup"><span data-stu-id="8d649-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="8d649-135">Они также применяют атрибуты к текущему пользователю надстройки.</span><span class="sxs-lookup"><span data-stu-id="8d649-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="8d649-136">В следующем примере кода добавляется ответ к первому примечанию в книге.</span><span class="sxs-lookup"><span data-stu-id="8d649-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="8d649-137">Редактирование комментариев</span><span class="sxs-lookup"><span data-stu-id="8d649-137">Edit comments</span></span>

<span data-ttu-id="8d649-138">Чтобы изменить примечание или ответ на примечание, настройте его свойство `Comment.content` или `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="8d649-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="8d649-139">Изменение ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="8d649-139">Edit comment replies</span></span>

<span data-ttu-id="8d649-140">Чтобы изменить ответ на комментарий, задайте его `CommentReply.content` свойство.</span><span class="sxs-lookup"><span data-stu-id="8d649-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="8d649-141">Удалять комментарии.</span><span class="sxs-lookup"><span data-stu-id="8d649-141">Delete comments</span></span>

<span data-ttu-id="8d649-142">Чтобы удалить комментарий, используйте `Comment.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="8d649-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="8d649-143">При удалении комментария также удаляются ответы, связанные с этим комментарием.</span><span class="sxs-lookup"><span data-stu-id="8d649-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="8d649-144">Удаление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="8d649-144">Delete comment replies</span></span>

<span data-ttu-id="8d649-145">Чтобы удалить ответ на комментарий, используйте `CommentReply.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="8d649-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="8d649-146">Разрешение потоков комментариев</span><span class="sxs-lookup"><span data-stu-id="8d649-146">Resolve comment threads</span></span>

<span data-ttu-id="8d649-147">Поток комментариев имеет настраиваемое логическое значение, `resolved` которое указывает, разрешено ли оно.</span><span class="sxs-lookup"><span data-stu-id="8d649-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="8d649-148">Значение означает, `true` что поток комментариев разрешается.</span><span class="sxs-lookup"><span data-stu-id="8d649-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="8d649-149">Значение означает, `false` что поток комментариев является либо новым, либо повторно открытым.</span><span class="sxs-lookup"><span data-stu-id="8d649-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="8d649-150">Ответы на комментарии имеют `resolved` свойство ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="8d649-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="8d649-151">Его значение всегда равно значению остальной части потока.</span><span class="sxs-lookup"><span data-stu-id="8d649-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="8d649-152">Метаданные Comment</span><span class="sxs-lookup"><span data-stu-id="8d649-152">Comment metadata</span></span>

<span data-ttu-id="8d649-153">Каждое примечание содержит метаданные о его создании, например автора и дату создания.</span><span class="sxs-lookup"><span data-stu-id="8d649-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="8d649-154">Автором примечаний, созданных вашей надстройкой, считается текущий пользователь.</span><span class="sxs-lookup"><span data-stu-id="8d649-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="8d649-155">В следующем примере показано, как отобразить электронную почту автора, имя автора и дату создания примечания в ячейке **A2**.</span><span class="sxs-lookup"><span data-stu-id="8d649-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="8d649-156">Метаданные ответа на комментарии</span><span class="sxs-lookup"><span data-stu-id="8d649-156">Comment reply metadata</span></span>

<span data-ttu-id="8d649-157">Ответы на комментарии хранят те же типы метаданных, что и исходный комментарий.</span><span class="sxs-lookup"><span data-stu-id="8d649-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="8d649-158">В приведенном ниже примере показано, как отобразить сообщение об авторе, имя автора и дату создания последнего ответа на комментарий в **ячейке A2**.</span><span class="sxs-lookup"><span data-stu-id="8d649-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="8d649-159">Упоминания</span><span class="sxs-lookup"><span data-stu-id="8d649-159">Mentions</span></span>

<span data-ttu-id="8d649-160">[Упоминания](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) используются для обозначения коллег в комментарии.</span><span class="sxs-lookup"><span data-stu-id="8d649-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="8d649-161">При этом уведомления отправляются с содержимым комментария.</span><span class="sxs-lookup"><span data-stu-id="8d649-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="8d649-162">Ваша надстройка может создавать эти упоминания от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="8d649-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="8d649-163">Комментарии с упоминанием следует создавать с помощью объектов [комментричконтент](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="8d649-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="8d649-164">Вызов `CommentCollection.add` с `CommentRichContent` указанием одного или нескольких упоминаний и указанием в `ContentType.mention` качестве `contentType` параметра.</span><span class="sxs-lookup"><span data-stu-id="8d649-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="8d649-165">`content`Строку также необходимо отформатировать, чтобы вставить упоминание в текст.</span><span class="sxs-lookup"><span data-stu-id="8d649-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="8d649-166">Формат для упоминания: `<at id="{replyIndex}">{mentionName}</at>` .</span><span class="sxs-lookup"><span data-stu-id="8d649-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="8d649-167">В настоящее время в качестве текста ссылки на упоминание можно использовать только точное имя упоминания.</span><span class="sxs-lookup"><span data-stu-id="8d649-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="8d649-168">Поддержка сокращенных версий имени будет добавлена позже.</span><span class="sxs-lookup"><span data-stu-id="8d649-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="8d649-169">В приведенном ниже примере показан комментарий с одним упоминанием.</span><span class="sxs-lookup"><span data-stu-id="8d649-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="8d649-170">См. также</span><span class="sxs-lookup"><span data-stu-id="8d649-170">See also</span></span>

- [<span data-ttu-id="8d649-171">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="8d649-171">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8d649-172">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="8d649-172">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="8d649-173">Вставка комментариев и заметок в Excel</span><span class="sxs-lookup"><span data-stu-id="8d649-173">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
