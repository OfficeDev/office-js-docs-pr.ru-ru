---
title: Работайте с комментариями с помощью API JavaScript для Excel
description: Сведения об использовании API для добавления, удаления и редактирования комментариев и потоков комментариев.
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: d6be0f07e0d3bb134385f0a08c20ce00da4de892
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327784"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="4e3fa-103">Работайте с комментариями с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="4e3fa-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="4e3fa-104">В этой статье описывается, как добавлять, читать, изменять и удалять комментарии в книге с помощью API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="4e3fa-105">Дополнительные сведения о функции комментариев можно узнать в статье [INSERT Comments and notess in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) .</span><span class="sxs-lookup"><span data-stu-id="4e3fa-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="4e3fa-106">В API JavaScript для Excel комментарий — это начальная заметка и связанное обсуждение.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-106">In the Excel JavaScript API, a comment is both the initial note and the connected threaded discussion.</span></span> <span data-ttu-id="4e3fa-107">Он привязан к отдельной ячейке.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-107">It is tied to an individual cell.</span></span> <span data-ttu-id="4e3fa-108">Любой пользователь, просматривающий книгу с достаточными разрешениями, может ответить на комментарий.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="4e3fa-109">Объект [comment](/javascript/api/excel/excel.comment) хранит эти ответы как объекты [комментрепли](/javascript/api/excel/excel.commentreply) .</span><span class="sxs-lookup"><span data-stu-id="4e3fa-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="4e3fa-110">Обратите внимание на то, что комментарий является потоком и что поток должен иметь специальную запись в качестве отправной точки.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![Комментарий Excel с пометкой "Comment" с двумя ответами, помеченными как "Comment. ответы [0]" и "Comment. ответы [1].](../images/excel-comments.png)

<span data-ttu-id="4e3fa-112">Комментарии в книге отслеживаются `Workbook.comments` свойством.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="4e3fa-113">Это касается примечаний, созданных пользователями, а также примечаний, созданных вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="4e3fa-114">Свойство `Workbook.comments` является объектом [CommentCollection](/javascript/api/excel/excel.commentcollection), содержащим коллекцию объектов [Comment](/javascript/api/excel/excel.comment).</span><span class="sxs-lookup"><span data-stu-id="4e3fa-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="4e3fa-115">Комментарии также доступны на уровне [листа](/javascript/api/excel/excel.worksheet) .</span><span class="sxs-lookup"><span data-stu-id="4e3fa-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="4e3fa-116">Примеры, приведенные в этой статье, работают с комментариями на уровне книги, но их можно легко изменить, `Worksheet.comments` чтобы использовать свойство.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="4e3fa-117">Добавление примечаний</span><span class="sxs-lookup"><span data-stu-id="4e3fa-117">Add comments</span></span>

<span data-ttu-id="4e3fa-118">Используйте `CommentCollection.add` метод, чтобы добавить комментарии в книгу.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="4e3fa-119">Этот метод занимает до трех параметров:</span><span class="sxs-lookup"><span data-stu-id="4e3fa-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="4e3fa-120">`cellAddress`: Ячейка, в которую добавляется комментарий.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="4e3fa-121">Это может быть объект String или [Range](/javascript/api/excel/excel.range) .</span><span class="sxs-lookup"><span data-stu-id="4e3fa-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="4e3fa-122">Диапазон должен быть одной ячейкой.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-122">The range must be a single cell.</span></span>
- <span data-ttu-id="4e3fa-123">`content`: Контент комментария.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-123">`content`: The comment's content.</span></span> <span data-ttu-id="4e3fa-124">Используйте строку для примечаний в виде обычного текста.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-124">Use a string for plain text comments.</span></span> <span data-ttu-id="4e3fa-125">Используйте объект [комментричконтент](/javascript/api/excel/excel.commentrichcontent) для комментариев с [упоминаниями](#mentions-online-only).</span><span class="sxs-lookup"><span data-stu-id="4e3fa-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions-online-only).</span></span> 
- <span data-ttu-id="4e3fa-126">`contentType`: Перечисление [ContentType](/javascript/api/excel/excel.contenttype) , определяющее тип контента.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="4e3fa-127">Значение по умолчанию — `ContentType.plain`.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="4e3fa-128">В следующем примере кода добавляется примечание в ячейку **A2**.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="4e3fa-129">Комментарии, добавленные надстройкой, добавляются к текущему пользователю этой надстройки.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="4e3fa-130">Добавление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="4e3fa-130">Add comment replies</span></span>

<span data-ttu-id="4e3fa-131">`Comment` Объект — это поток комментариев, который содержит ноль или больше ответов.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="4e3fa-132">Объекты `Comment` содержат свойство `replies`, являющееся коллекцией [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection), содержащей объекты [CommentReply](/javascript/api/excel/excel.commentreply).</span><span class="sxs-lookup"><span data-stu-id="4e3fa-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="4e3fa-133">Чтобы добавить ответ на примечание, используйте метод `CommentReplyCollection.add`, передающий текст ответа.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="4e3fa-134">Ответы отображаются в порядке их добавления.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="4e3fa-135">Они также применяют атрибуты к текущему пользователю надстройки.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="4e3fa-136">В следующем примере кода добавляется ответ к первому примечанию в книге.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="4e3fa-137">Редактирование комментариев</span><span class="sxs-lookup"><span data-stu-id="4e3fa-137">Edit comments</span></span>

<span data-ttu-id="4e3fa-138">Чтобы изменить примечание или ответ на примечание, настройте его свойство `Comment.content` или `CommentReply.content`.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="4e3fa-139">Изменение ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="4e3fa-139">Edit comment replies</span></span>

<span data-ttu-id="4e3fa-140">Чтобы изменить ответ на комментарий, задайте его `CommentReply.content` свойство.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="4e3fa-141">Удалять комментарии.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-141">Delete comments</span></span>

<span data-ttu-id="4e3fa-142">Чтобы удалить комментарий, `Comment.delete` используйте метод.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="4e3fa-143">При удалении комментария также удаляются ответы, связанные с этим комментарием.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="4e3fa-144">Удаление ответов на комментарии</span><span class="sxs-lookup"><span data-stu-id="4e3fa-144">Delete comment replies</span></span>

<span data-ttu-id="4e3fa-145">Чтобы удалить ответ на комментарий, используйте `CommentReply.delete` метод.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a><span data-ttu-id="4e3fa-146">Разрешение потоков комментариев ([Предварительная версия](../reference/requirement-sets/excel-preview-apis.md))</span><span class="sxs-lookup"><span data-stu-id="4e3fa-146">Resolve comment threads ([preview](../reference/requirement-sets/excel-preview-apis.md))</span></span> 

<span data-ttu-id="4e3fa-147">Поток комментариев имеет настраиваемое логическое значение `resolved`, которое указывает, разрешено ли оно.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="4e3fa-148">Значение `true` означает, что поток комментариев разрешается.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="4e3fa-149">Значение `false` означает, что поток комментариев является либо новым, либо повторно открытым.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="4e3fa-150">Ответы на комментарии имеют свойство `resolved` ReadOnly.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="4e3fa-151">Его значение всегда равно значению остальной части потока.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="4e3fa-152">Метаданные Comment</span><span class="sxs-lookup"><span data-stu-id="4e3fa-152">Comment metadata</span></span>

<span data-ttu-id="4e3fa-153">Каждое примечание содержит метаданные о его создании, например автора и дату создания.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="4e3fa-154">Автором примечаний, созданных вашей надстройкой, считается текущий пользователь.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="4e3fa-155">В следующем примере показано, как отобразить электронную почту автора, имя автора и дату создания примечания в ячейке **A2**.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="4e3fa-156">Метаданные ответа на комментарии</span><span class="sxs-lookup"><span data-stu-id="4e3fa-156">Comment reply metadata</span></span>

<span data-ttu-id="4e3fa-157">Ответы на комментарии хранят те же типы метаданных, что и исходный комментарий.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="4e3fa-158">В приведенном ниже примере показано, как отобразить сообщение об авторе, имя автора и дату создания последнего ответа на комментарий в **ячейке A2**.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions-online-only"></a><span data-ttu-id="4e3fa-159">Упоминания ([только в Интернете](../reference/requirement-sets/excel-api-online-requirement-set.md))</span><span class="sxs-lookup"><span data-stu-id="4e3fa-159">Mentions ([online-only](../reference/requirement-sets/excel-api-online-requirement-set.md))</span></span> 

> [!NOTE]
> <span data-ttu-id="4e3fa-160">API упомянутых комментариев в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-160">The comment mention APIs are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> <span data-ttu-id="4e3fa-161">Упоминание комментариев в настоящее время поддерживается только для Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-161">Comment mentions are currently only supported for Excel on the web.</span></span>

<span data-ttu-id="4e3fa-162">[Упоминания](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) используются для обозначения коллег в комментарии.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-162">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="4e3fa-163">При этом уведомления отправляются с содержимым комментария.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-163">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="4e3fa-164">Ваша надстройка может создавать эти упоминания от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-164">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="4e3fa-165">Комментарии с упоминанием следует создавать с помощью объектов [комментричконтент](/javascript/api/excel/excel.commentrichcontent) .</span><span class="sxs-lookup"><span data-stu-id="4e3fa-165">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="4e3fa-166">Вызов `CommentCollection.add` с `CommentRichContent` указанием одного или нескольких упоминаний и указанием `ContentType.mention` в `contentType` качестве параметра.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-166">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="4e3fa-167">`content` Строку также необходимо отформатировать, чтобы вставить упоминание в текст.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-167">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="4e3fa-168">Формат для упоминания: `<at id="{replyIndex}">{mentionName}</at>`.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-168">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> <span data-ttu-id="4e3fa-169">НОТЕ В настоящее время в качестве текста ссылки на упоминание можно использовать только точное имя упоминания.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-169">[NOTE] Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="4e3fa-170">Поддержка сокращенных версий имени будет добавлена позже.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-170">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="4e3fa-171">В приведенном ниже примере показан комментарий с одним упоминанием.</span><span class="sxs-lookup"><span data-stu-id="4e3fa-171">The following example shows a comment with a single mention.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4e3fa-172">См. также</span><span class="sxs-lookup"><span data-stu-id="4e3fa-172">See also</span></span>

- [<span data-ttu-id="4e3fa-173">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="4e3fa-173">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4e3fa-174">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="4e3fa-174">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="4e3fa-175">Вставка комментариев и заметок в Excel</span><span class="sxs-lookup"><span data-stu-id="4e3fa-175">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
