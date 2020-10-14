---
title: Работайте с комментариями с помощью API JavaScript для Excel
description: Сведения об использовании API для добавления, удаления и редактирования комментариев и потоков комментариев.
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 85312cbd92aa6c9d0f82fd167e8a372c2eff8c85
ms.sourcegitcommit: b50eebd303adcc22eb86e65756ce7e9a82f41a57
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/14/2020
ms.locfileid: "48456554"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="449cb-103">Работайте с комментариями с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="449cb-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="449cb-104">В этой статье описывается, как добавлять, читать, изменять и удалять комментарии в книге с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="449cb-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="449cb-105">Дополнительные сведения о функции комментариев можно узнать в статье [INSERT Comments and notess in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="449cb-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="449cb-106">В API JavaScript для Excel комментарий включает один начальный комментарий и подключенное обсуждение.</span><span class="sxs-lookup"><span data-stu-id="449cb-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="449cb-107">Он привязан к отдельной ячейке.</span><span class="sxs-lookup"><span data-stu-id="449cb-107">It is tied to an individual cell.</span></span> <span data-ttu-id="449cb-108">Любой пользователь, просматривающий книгу с достаточными разрешениями, может ответить на комментарий.</span><span class="sxs-lookup"><span data-stu-id="449cb-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="449cb-109">Объект [comment](/javascript/api/excel/excel.comment) хранит эти ответы как объекты [комментрепли](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="449cb-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="449cb-110">Обратите внимание на то, что комментарий является потоком и что поток должен иметь специальную запись в качестве отправной точки.</span><span class="sxs-lookup"><span data-stu-id="449cb-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Комментарий Excel с пометкой "Comment" с двумя ответами, помеченными как "Comment. ответы [0]" и "Comment. ответы [1].](../images/excel-comments.png)

<span data-ttu-id="449cb-112">Комментарии в книге отслеживаются `Workbook.comments` свойством.</span><span class="sxs-lookup"><span data-stu-id="449cb-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="449cb-113">Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="449cb-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="449cb-114">Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="449cb-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="449cb-115">Комментарии также доступны на уровне [листа](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="449cb-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="449cb-116">Примеры, приведенные в этой статье, работают с комментариями на уровне книги, но их можно легко изменить, чтобы использовать `Worksheet.comments` свойство.</span><span class="sxs-lookup"><span data-stu-id="449cb-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="449cb-117">Добавление примечаний</span><span class="sxs-lookup"><span data-stu-id="449cb-117">Add comments</span></span>

<span data-ttu-id="449cb-118">Используйте `CommentCollection.add` метод, чтобы добавить комментарии в книгу.</span><span class="sxs-lookup"><span data-stu-id="449cb-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="449cb-119">Этот метод занимает до трех параметров:</span><span class="sxs-lookup"><span data-stu-id="449cb-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="449cb-120">`cellAddress`: Ячейка, в которую добавляется комментарий.</span><span class="sxs-lookup"><span data-stu-id="449cb-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="449cb-121">Это может быть объект String или [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="449cb-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="449cb-122">Диапазон должен быть одной ячейкой.</span><span class="sxs-lookup"><span data-stu-id="449cb-122">The range must be a single cell.</span></span>
- <span data-ttu-id="449cb-123">`content`: Контент комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-123">`content`: The comment's content.</span></span> <span data-ttu-id="449cb-124">Используйте строку для примечаний в виде обычного текста.</span><span class="sxs-lookup"><span data-stu-id="449cb-124">Use a string for plain text comments.</span></span> <span data-ttu-id="449cb-125">Используйте объект [комментричконтент](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями](#mentions).</span><span class="sxs-lookup"><span data-stu-id="449cb-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="449cb-126">`contentType`: Перечисление [ContentType](/javascript/api/excel/excel.contenttype) , определяющее тип контента.</span><span class="sxs-lookup"><span data-stu-id="449cb-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="449cb-127">Значение по умолчанию — `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="449cb-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="449cb-128">В следующем примере кода добавляется примечание в ячейку **A2**.</span><span class="sxs-lookup"><span data-stu-id="449cb-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="449cb-129">Комментарии, добавленные надстройкой, добавляются к текущему пользователю этой надстройки.</span><span class="sxs-lookup"><span data-stu-id="449cb-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="449cb-130">Добавление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="449cb-130">Add comment replies</span></span>

<span data-ttu-id="449cb-131">`Comment`Объект — это поток комментариев, который содержит ноль или больше ответов.</span><span class="sxs-lookup"><span data-stu-id="449cb-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="449cb-132">Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="449cb-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="449cb-133">Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа.</span><span class="sxs-lookup"><span data-stu-id="449cb-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="449cb-134">Ответы отображаются в порядке их добавления.</span><span class="sxs-lookup"><span data-stu-id="449cb-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="449cb-135">Они также применяют атрибуты к текущему пользователю надстройки.</span><span class="sxs-lookup"><span data-stu-id="449cb-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="449cb-136">В следующем примере кода добавляется ответ к первому примечанию в книге.</span><span class="sxs-lookup"><span data-stu-id="449cb-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="449cb-137">Редактирование комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-137">Edit comments</span></span>

<span data-ttu-id="449cb-138">Чтобы изменить примечание или ответ на примечание, настройте его свойство `Comment.content` или `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="449cb-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="449cb-139">Изменение ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="449cb-139">Edit comment replies</span></span>

<span data-ttu-id="449cb-140">Чтобы изменить ответ на комментарий, задайте его `CommentReply.content` свойство.</span><span class="sxs-lookup"><span data-stu-id="449cb-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="449cb-141">Удалять комментарии.</span><span class="sxs-lookup"><span data-stu-id="449cb-141">Delete comments</span></span>

<span data-ttu-id="449cb-142">Чтобы удалить комментарий, используйте `Comment.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="449cb-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="449cb-143">При удалении комментария также удаляются ответы, связанные с этим комментарием.</span><span class="sxs-lookup"><span data-stu-id="449cb-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="449cb-144">Удаление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="449cb-144">Delete comment replies</span></span>

<span data-ttu-id="449cb-145">Чтобы удалить ответ на комментарий, используйте `CommentReply.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="449cb-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="449cb-146">Разрешение потоков комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-146">Resolve comment threads</span></span>

<span data-ttu-id="449cb-147">Поток комментариев имеет настраиваемое логическое значение, `resolved` которое указывает, разрешено ли оно.</span><span class="sxs-lookup"><span data-stu-id="449cb-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="449cb-148">Значение означает, `true` что поток комментариев разрешается.</span><span class="sxs-lookup"><span data-stu-id="449cb-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="449cb-149">Значение означает, `false` что поток комментариев является либо новым, либо повторно открытым.</span><span class="sxs-lookup"><span data-stu-id="449cb-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="449cb-150">Ответы на комментарии имеют `resolved` свойство ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="449cb-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="449cb-151">Его значение всегда равно значению остальной части потока.</span><span class="sxs-lookup"><span data-stu-id="449cb-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="449cb-152">Метаданные Comment</span><span class="sxs-lookup"><span data-stu-id="449cb-152">Comment metadata</span></span>

<span data-ttu-id="449cb-153">Каждое примечание содержит метаданные о его создании, например автора и дату создания.</span><span class="sxs-lookup"><span data-stu-id="449cb-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="449cb-154">Автором примечаний, созданных вашей надстройкой, считается текущий пользователь.</span><span class="sxs-lookup"><span data-stu-id="449cb-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="449cb-155">В следующем примере показано, как отобразить электронную почту автора, имя автора и дату создания примечания в ячейке **A2**.</span><span class="sxs-lookup"><span data-stu-id="449cb-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="449cb-156">Метаданные ответа на комментарии</span><span class="sxs-lookup"><span data-stu-id="449cb-156">Comment reply metadata</span></span>

<span data-ttu-id="449cb-157">Ответы на комментарии хранят те же типы метаданных, что и исходный комментарий.</span><span class="sxs-lookup"><span data-stu-id="449cb-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="449cb-158">В приведенном ниже примере показано, как отобразить сообщение об авторе, имя автора и дату создания последнего ответа на комментарий в **ячейке A2**.</span><span class="sxs-lookup"><span data-stu-id="449cb-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="449cb-159">Упоминания</span><span class="sxs-lookup"><span data-stu-id="449cb-159">Mentions</span></span>

<span data-ttu-id="449cb-160">[Упоминания](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) используются для обозначения коллег в комментарии.</span><span class="sxs-lookup"><span data-stu-id="449cb-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="449cb-161">При этом уведомления отправляются с содержимым комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="449cb-162">Ваша надстройка может создавать эти упоминания от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="449cb-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="449cb-163">Комментарии с упоминанием следует создавать с помощью объектов [комментричконтент](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="449cb-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="449cb-164">Вызов `CommentCollection.add` с `CommentRichContent` указанием одного или нескольких упоминаний и указанием в `ContentType.mention` качестве `contentType` параметра.</span><span class="sxs-lookup"><span data-stu-id="449cb-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="449cb-165">`content`Строку также необходимо отформатировать, чтобы вставить упоминание в текст.</span><span class="sxs-lookup"><span data-stu-id="449cb-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="449cb-166">Формат для упоминания: `<at id="{replyIndex}">{mentionName}</at>` .</span><span class="sxs-lookup"><span data-stu-id="449cb-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="449cb-167">В настоящее время в качестве текста ссылки на упоминание можно использовать только точное имя упоминания.</span><span class="sxs-lookup"><span data-stu-id="449cb-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="449cb-168">Поддержка сокращенных версий имени будет добавлена позже.</span><span class="sxs-lookup"><span data-stu-id="449cb-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="449cb-169">В приведенном ниже примере показан комментарий с одним упоминанием.</span><span class="sxs-lookup"><span data-stu-id="449cb-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="comment-events"></a><span data-ttu-id="449cb-170">События комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-170">Comment events</span></span>

<span data-ttu-id="449cb-171">Надстройка может прослушивать добавленные комментарии, изменения и удаления.</span><span class="sxs-lookup"><span data-stu-id="449cb-171">Your add-in can listen for comment additions, changes, and deletions.</span></span> <span data-ttu-id="449cb-172">[События с комментариями](/javascript/api/excel/excel.commentcollection#event-details) возникают для `CommentCollection` объекта.</span><span class="sxs-lookup"><span data-stu-id="449cb-172">[Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.</span></span> <span data-ttu-id="449cb-173">Для прослушивания событий комментариев Зарегистрируйте `onAdded` `onChanged` обработчик событий, или `onDeleted` комментарий.</span><span class="sxs-lookup"><span data-stu-id="449cb-173">To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler.</span></span> <span data-ttu-id="449cb-174">При обнаружении события комментария используйте этот обработчик событий для получения данных о добавленных, измененных или удаленных комментариях.</span><span class="sxs-lookup"><span data-stu-id="449cb-174">When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment.</span></span> <span data-ttu-id="449cb-175">`onChanged`Кроме того, событие обрабатывает добавления, изменения и удаления ответа на комментарии.</span><span class="sxs-lookup"><span data-stu-id="449cb-175">The `onChanged` event also handles comment reply additions, changes, and deletions.</span></span> 

<span data-ttu-id="449cb-176">Каждое событие Comment инициируется только один раз при одновременном выполнении нескольких добавлений, изменений или удалений.</span><span class="sxs-lookup"><span data-stu-id="449cb-176">Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time.</span></span> <span data-ttu-id="449cb-177">Все объекты [комментаддедевентаргс](/javascript/api/excel/excel.commentaddedeventargs), [комментчанжедевентаргс](/javascript/api/excel/excel.commentchangedeventarg)и [комментделетедевентаргс](/javascript/api/excel/excel.commentdeletedeventargs) содержат массивы идентификаторов комментариев для сопоставления действий события с коллекциями комментариев.</span><span class="sxs-lookup"><span data-stu-id="449cb-177">All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.</span></span>

<span data-ttu-id="449cb-178">В статье [Работа с событиями с помощью API JavaScript для Excel](excel-add-ins-events.md) можно получить дополнительные сведения о регистрации обработчиков событий, обработке событий и удалении обработчиков событий.</span><span class="sxs-lookup"><span data-stu-id="449cb-178">See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.</span></span> 

### <a name="comment-addition-events"></a><span data-ttu-id="449cb-179">События добавления комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-179">Comment addition events</span></span> 
<span data-ttu-id="449cb-180">`onAdded`Событие инициируется, когда в коллекцию комментариев добавляется один или несколько новых комментариев.</span><span class="sxs-lookup"><span data-stu-id="449cb-180">The `onAdded` event is triggered when one or more new comments are added to the comment collection.</span></span> <span data-ttu-id="449cb-181">Это событие *не* инициируется при добавлении ответов в поток комментариев (просмотрите [события изменения комментария](#comment-change-events) , чтобы узнать о событиях ответа на комментарии).</span><span class="sxs-lookup"><span data-stu-id="449cb-181">This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).</span></span>

<span data-ttu-id="449cb-182">В следующем примере показано, как зарегистрировать `onAdded` обработчик событий и затем использовать `CommentAddedEventArgs` объект для получения `commentDetails` массива добавленного комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-182">The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.</span></span>

> [!NOTE]
> <span data-ttu-id="449cb-183">Этот пример работает только при добавлении одного комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-183">This sample only works when a single comment is added.</span></span> 

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

### <a name="comment-change-events"></a><span data-ttu-id="449cb-184">События изменения комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-184">Comment change events</span></span> 
<span data-ttu-id="449cb-185">`onChanged`Событие Comment запускается в приведенных ниже сценариях.</span><span class="sxs-lookup"><span data-stu-id="449cb-185">The `onChanged` comment event is triggered in the following scenarios.</span></span>

- <span data-ttu-id="449cb-186">Обновляется контент комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-186">A comment's content is updated.</span></span>
- <span data-ttu-id="449cb-187">Поток комментариев разрешается.</span><span class="sxs-lookup"><span data-stu-id="449cb-187">A comment thread is resolved.</span></span>
- <span data-ttu-id="449cb-188">Поток комментариев повторно открыт.</span><span class="sxs-lookup"><span data-stu-id="449cb-188">A comment thread is reopened.</span></span>
- <span data-ttu-id="449cb-189">В цепочку комментариев добавляется ответ.</span><span class="sxs-lookup"><span data-stu-id="449cb-189">A reply is added to a comment thread.</span></span>
- <span data-ttu-id="449cb-190">Ответ обновляется в цепочке комментариев.</span><span class="sxs-lookup"><span data-stu-id="449cb-190">A reply is updated in a comment thread.</span></span>
- <span data-ttu-id="449cb-191">В цепочке комментариев удаляется ответ.</span><span class="sxs-lookup"><span data-stu-id="449cb-191">A reply is deleted in a comment thread.</span></span>

<span data-ttu-id="449cb-192">В следующем примере показано, как зарегистрировать `onChanged` обработчик событий и затем использовать `CommentChangedEventArgs` объект для получения `commentDetails` массива измененного комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-192">The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.</span></span>

> [!NOTE]
> <span data-ttu-id="449cb-193">Этот пример работает только при изменении одного комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-193">This sample only works when a single comment is changed.</span></span> 

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

### <a name="comment-deletion-events"></a><span data-ttu-id="449cb-194">События удаления комментариев</span><span class="sxs-lookup"><span data-stu-id="449cb-194">Comment deletion events</span></span>
<span data-ttu-id="449cb-195">`onDeleted`Событие инициируется при удалении комментария из коллекции комментариев.</span><span class="sxs-lookup"><span data-stu-id="449cb-195">The `onDeleted` event is triggered when a comment is deleted from the comment collection.</span></span> <span data-ttu-id="449cb-196">После удаления комментария его метаданные больше не будут доступны.</span><span class="sxs-lookup"><span data-stu-id="449cb-196">Once a comment has been deleted, its metadata is no longer available.</span></span> <span data-ttu-id="449cb-197">Объект [комментделетедевентаргс](/javascript/api/excel/excel.commentdeletedeventargs) предоставляет идентификаторы комментариев, если ваша надстройка управляет отдельными комментариями.</span><span class="sxs-lookup"><span data-stu-id="449cb-197">The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.</span></span>

<span data-ttu-id="449cb-198">В приведенном ниже примере показано, как зарегистрировать `onDeleted` обработчик событий и затем использовать `CommentDeletedEventArgs` объект для получения `commentDetails` массива удаляемого комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-198">The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.</span></span>

> [!NOTE]
> <span data-ttu-id="449cb-199">Этот пример работает только при удалении одного комментария.</span><span class="sxs-lookup"><span data-stu-id="449cb-199">This sample only works when a single comment is deleted.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="449cb-200">См. также</span><span class="sxs-lookup"><span data-stu-id="449cb-200">See also</span></span>

- [<span data-ttu-id="449cb-201">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="449cb-201">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="449cb-202">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="449cb-202">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="449cb-203">Работа с событиями при помощи API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="449cb-203">Work with Events using the Excel JavaScript API</span></span>](excel-add-ins-events.md)
- [<span data-ttu-id="449cb-204">Вставка комментариев и заметок в Excel</span><span class="sxs-lookup"><span data-stu-id="449cb-204">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
