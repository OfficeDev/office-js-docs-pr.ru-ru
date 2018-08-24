---
title: Обзор надстроек Word
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 90b6ad29b4d1fe4ecffd3a1e57d76d16752ee77a
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925432"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="7117e-102">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="7117e-102">Word add-ins overview</span></span>

<span data-ttu-id="7117e-p101">Хотите создать решение для автоматического составления документов или привязки и доступа к данным в документе Word из других источников? Чтобы расширить возможности клиентов Word на компьютере с Windows, Mac или в облаке, используйте платформу надстроек Office, которая включает API JavaScript для Word и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="7117e-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="7117e-p102">На [платформе надстроек Office](../overview/office-add-ins.md) можно разрабатывать не только надстройки Word. Используя команды надстроек, вы можете расширять интерфейс Word и запускать области задач, которые выполняют сценарий JavaScript, взаимодействующий с содержимым документа. Любой код, который работает в браузере, будет работать в надстройке Word. Надстройки, взаимодействующие с содержимым документа Word, создают запросы на совершение действий с объектами Word и синхронизацию состояния этих объектов.</span><span class="sxs-lookup"><span data-stu-id="7117e-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="7117e-p103">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource, она должна соответствовать [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и [статье о доступности надстроек Office в ведущих приложениях](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="7117e-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="7117e-113">Ниже показан пример надстройки Word, работающей в области задач.</span><span class="sxs-lookup"><span data-stu-id="7117e-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="7117e-114">*Рис. 1. Надстройка, работающая в области задач Word*</span><span class="sxs-lookup"><span data-stu-id="7117e-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Надстройка, работающая в области задач Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="7117e-p104">Надстройка Word может (1) отправлять запросы в документ Word и (2) обновлять, удалять или перемещать абзац, используя JavaScript для доступа к объекту paragraph. Например, в приведенном ниже коде показано, как добавить в абзац новое предложение.</span><span class="sxs-lookup"><span data-stu-id="7117e-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="7117e-p105">Для размещения надстройки Word можно использовать любой веб-сервер, в частности ASP.NET, NodeJS и Python. Используйте любимую клиентскую платформу — Ember, Backbone, Angular, React —для разработки своего решения; или продолжайте работу с VanillaJS. Для [аутентификации](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) и размещения приложения можно использовать Azure.</span><span class="sxs-lookup"><span data-stu-id="7117e-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="7117e-p106">API JavaScript для Word предоставляют приложению доступ к объектам и метаданным документа Word. С помощью этих API можно создавать надстройки, предназначенные для:</span><span class="sxs-lookup"><span data-stu-id="7117e-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="7117e-122">Word 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7117e-122">Word 2013 for Windows</span></span>
* <span data-ttu-id="7117e-123">Word 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7117e-123">Word 2016 for Windows</span></span>
* <span data-ttu-id="7117e-124">Word Online</span><span class="sxs-lookup"><span data-stu-id="7117e-124">Word Online</span></span>
* <span data-ttu-id="7117e-125">Word 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="7117e-125">Word 2016 for Mac</span></span>
* <span data-ttu-id="7117e-126">Word для iOS</span><span class="sxs-lookup"><span data-stu-id="7117e-126">Word for iOS</span></span>

<span data-ttu-id="7117e-p107">Написанные вами надстройки будут работать во всех версиях Word на различных платформах. Дополнительные сведения см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="7117e-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="7117e-129">API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="7117e-129">JavaScript APIs for Word</span></span>

<span data-ttu-id="7117e-p108">Для взаимодействия с объектами и метаданными в документе Word можно использовать два набора API JavaScript. Первый — [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) — появился в Office 2013. Это общий API — многие объекты можно использовать в надстройках, размещенных в двух или более клиентах Office. В этом API широко используются обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="7117e-p108">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word), which was introduced in Office 2013. This is a shared API -- many of the objects can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.</span></span> 

<span data-ttu-id="7117e-p109">Второй — [API JavaScript для Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). Это строго типизированная объектная модель, с помощью которой можно создавать надстройки Word, предназначенные для Word 2016 для Mac и Windows. Эта объектная модель использует обещания и предоставляет доступ к объектам Word, в частности [Body](https://dev.office.com/reference/add-ins/word/body), [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol), [InlinePicture](https://dev.office.com/reference/add-ins/word/inlinepicture) и [Paragraph](https://dev.office.com/reference/add-ins/word/paragraph). API JavaScript для Word включает определения TypeScript и файлы vsdoc, чтобы вы могли получать подсказки кода в своей интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="7117e-p109">The second is the [Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](https://dev.office.com/reference/add-ins/word/body), [content controls](https://dev.office.com/reference/add-ins/word/contentcontrol), [inline pictures](https://dev.office.com/reference/add-ins/word/inlinepicture), and [paragraphs](https://dev.office.com/reference/add-ins/word/paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="7117e-p110">В настоящее время все клиенты Word поддерживают общий API JavaScript для Office, а большинство из них поддерживают и API JavaScript для Word. Дополнительные сведения о поддерживаемых клиентах см. в [справочнике по API](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).</span><span class="sxs-lookup"><span data-stu-id="7117e-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="7117e-p111">Рекомендуем начать с API JavaScript для Word, так как с объектной моделью проще работать. Используйте API JavaScript для Word, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="7117e-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="7117e-142">получить доступ к объектам в документе Word.</span><span class="sxs-lookup"><span data-stu-id="7117e-142">Access the objects in a Word document.</span></span>

<span data-ttu-id="7117e-143">Используйте общий API JavaScript для Office, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="7117e-143">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="7117e-144">создать надстройки для Word 2013;</span><span class="sxs-lookup"><span data-stu-id="7117e-144">Target Word 2013.</span></span>
* <span data-ttu-id="7117e-145">выполнить начальные действия для приложения;</span><span class="sxs-lookup"><span data-stu-id="7117e-145">Perform initial actions for the application.</span></span>
* <span data-ttu-id="7117e-146">проверить поддерживаемый набор требований;</span><span class="sxs-lookup"><span data-stu-id="7117e-146">Check the supported requirement set.</span></span>
* <span data-ttu-id="7117e-147">получить доступ к метаданным документа, его параметрам и сведениям о среде;</span><span class="sxs-lookup"><span data-stu-id="7117e-147">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="7117e-148">создать привязку к разделам документа и записать события;</span><span class="sxs-lookup"><span data-stu-id="7117e-148">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="7117e-149">использовать пользовательские XML-части;</span><span class="sxs-lookup"><span data-stu-id="7117e-149">Use custom XML parts.</span></span>
* <span data-ttu-id="7117e-150">открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7117e-150">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7117e-151">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7117e-151">Next steps</span></span>

<span data-ttu-id="7117e-p112">Готовы [создать свою первую надстройку Word](word-add-ins.md)? Вы также можете воспользоваться нашим интерактивным [руководством по началу работы](http://dev.office.com/getting-started/addins?product=Word). Используйте [манифест надстройки](../develop/add-in-manifests.md), чтобы указать ведущее приложение, имя, разрешения и другие сведения.</span><span class="sxs-lookup"><span data-stu-id="7117e-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](http://dev.office.com/getting-started/addins?product=Word). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="7117e-156">Чтобы узнать больше о том, как создать качественную и привлекательную надстройку Word, см. [руководство по разработке](../design/add-in-design.md) и [рекомендации](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="7117e-156">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="7117e-157">После разработки надстройку можно [опубликовать](../publish/publish.md) в сетевой папке, каталоге приложений или AppSource.</span><span class="sxs-lookup"><span data-stu-id="7117e-157">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="whats-coming-up-for-word-add-ins"></a><span data-ttu-id="7117e-158">Над чем мы работаем?</span><span class="sxs-lookup"><span data-stu-id="7117e-158">What's coming up for Word add-ins?</span></span>

<span data-ttu-id="7117e-p113">Мы публикуем новые API для надстроек Word на странице [Открытые спецификации API](https://dev.office.com/reference/add-ins/openspec), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями API JavaScript для Word мы работаем, и поделитесь своим мнением о проектируемых спецификациях.</span><span class="sxs-lookup"><span data-stu-id="7117e-p113">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [API open specifications](https://dev.office.com/reference/add-ins/openspec) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="7117e-161">См. также</span><span class="sxs-lookup"><span data-stu-id="7117e-161">See also</span></span>

* [<span data-ttu-id="7117e-162">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7117e-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="7117e-163">Справочные материалы по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="7117e-163">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

