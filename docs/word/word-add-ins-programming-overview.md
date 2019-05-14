---
title: Обзор надстроек Word
description: ''
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 143f5b431aff2133c084b6d0f9c390562116dd4e
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952175"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="2a82f-102">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="2a82f-102">Word add-ins overview</span></span>

<span data-ttu-id="2a82f-p101">Хотите создать решение для автоматического составления документов или привязки и доступа к данным в документе Word из других источников? Чтобы расширить возможности клиентов Word на компьютере с Windows, Mac или в облаке, используйте платформу надстроек Office, которая включает API JavaScript для Word и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="2a82f-p102">На [платформе надстроек Office](../overview/office-add-ins.md) можно разрабатывать не только надстройки Word. Используя команды надстроек, вы можете расширять интерфейс Word и запускать области задач, которые выполняют сценарий JavaScript, взаимодействующий с содержимым документа. Любой код, который работает в браузере, будет работать в надстройке Word. Надстройки, взаимодействующие с содержимым документа Word, создают запросы на совершение действий с объектами Word и синхронизацию состояния этих объектов.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="2a82f-p103">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource, она должна соответствовать [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и [статье о доступности надстроек Office в ведущих приложениях](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="2a82f-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="2a82f-113">Ниже показан пример надстройки Word, работающей в области задач.</span><span class="sxs-lookup"><span data-stu-id="2a82f-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="2a82f-114">*Рис. 1. Надстройка, работающая в области задач Word*</span><span class="sxs-lookup"><span data-stu-id="2a82f-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Надстройка, работающая в области задач Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="2a82f-p104">Надстройка Word может (1) отправлять запросы в документ Word и (2) обновлять, удалять или перемещать абзац, используя JavaScript для доступа к объекту paragraph. Например, в приведенном ниже коде показано, как добавить в абзац новое предложение.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

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

<span data-ttu-id="2a82f-p105">Для размещения надстройки Word можно использовать любой веб-сервер, в частности ASP.NET, NodeJS и Python. Используйте любимую клиентскую платформу — Ember, Backbone, Angular, React —для разработки своего решения; или продолжайте работу с VanillaJS. Для [аутентификации](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) и размещения приложения можно использовать Azure.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="2a82f-p106">API JavaScript для Word предоставляют приложению доступ к объектам и метаданным документа Word. С помощью этих API можно создавать надстройки, предназначенные для:</span><span class="sxs-lookup"><span data-stu-id="2a82f-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="2a82f-122">Word 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="2a82f-122">Word 2013 or later for Windows</span></span>
* <span data-ttu-id="2a82f-123">Word Online</span><span class="sxs-lookup"><span data-stu-id="2a82f-123">Word Online</span></span>
* <span data-ttu-id="2a82f-124">Word 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="2a82f-124">Word 2016 or later for Mac</span></span>
* <span data-ttu-id="2a82f-125">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="2a82f-125">Word for iPad</span></span>

<span data-ttu-id="2a82f-p107">Написанные вами надстройки будут работать во всех версиях Word на различных платформах. Дополнительные сведения см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="2a82f-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="2a82f-128">API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2a82f-128">JavaScript APIs for Word</span></span>

<span data-ttu-id="2a82f-129">Для взаимодействия с объектами и метаданными в документе Word можно использовать два набора API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2a82f-129">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="2a82f-130">Первый — [общий API](../reference/javascript-api-for-office.md), представленный в Office 2013.</span><span class="sxs-lookup"><span data-stu-id="2a82f-130">The first is the [Common API](../reference/javascript-api-for-office.md), which was introduced in Office 2013.</span></span> <span data-ttu-id="2a82f-131">Многие объекты общего API можно использовать в надстройках, размещенных в двух или более клиентах Office.</span><span class="sxs-lookup"><span data-stu-id="2a82f-131">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="2a82f-132">В этом API широко используются обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="2a82f-132">This API uses callbacks extensively.</span></span>

<span data-ttu-id="2a82f-p109">Второй — [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md). Это строго типизированная объектная модель, с помощью которой можно создавать надстройки Word, предназначенные для Word 2016 для Mac и Windows. Эта объектная модель использует обещания и предоставляет доступ к объектам Word, в частности [Body](/javascript/api/word/word.body), [ContentControl](/javascript/api/word/word.contentcontrol), [InlinePicture](/javascript/api/word/word.inlinepicture) и [Paragraph](/javascript/api/word/word.paragraph). API JavaScript для Word включает определения TypeScript и файлы vsdoc, чтобы вы могли получать подсказки кода в своей интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p109">The second is the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="2a82f-p110">В настоящее время все клиенты Word поддерживают общий API JavaScript для Office, а большинство из них поддерживают и API JavaScript для Word. Дополнительные сведения о поддерживаемых клиентах см. в [справочнике по API](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span><span class="sxs-lookup"><span data-stu-id="2a82f-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](/office/dev/add-ins/reference/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="2a82f-p111">Рекомендуем начать с API JavaScript для Word, так как с объектной моделью проще работать. Используйте API JavaScript для Word, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="2a82f-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="2a82f-141">получить доступ к объектам в документе Word.</span><span class="sxs-lookup"><span data-stu-id="2a82f-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="2a82f-142">Используйте общий API JavaScript для Office, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="2a82f-142">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="2a82f-143">создать надстройки для Word 2013;</span><span class="sxs-lookup"><span data-stu-id="2a82f-143">Target Word 2013.</span></span>
* <span data-ttu-id="2a82f-144">выполнить начальные действия для приложения;</span><span class="sxs-lookup"><span data-stu-id="2a82f-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="2a82f-145">проверить поддерживаемый набор требований;</span><span class="sxs-lookup"><span data-stu-id="2a82f-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="2a82f-146">получить доступ к метаданным документа, его параметрам и сведениям о среде;</span><span class="sxs-lookup"><span data-stu-id="2a82f-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="2a82f-147">создать привязку к разделам документа и записать события;</span><span class="sxs-lookup"><span data-stu-id="2a82f-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="2a82f-148">использовать пользовательские XML-части;</span><span class="sxs-lookup"><span data-stu-id="2a82f-148">Use custom XML parts.</span></span>
* <span data-ttu-id="2a82f-149">открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="2a82f-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="2a82f-150">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="2a82f-150">Next steps</span></span>

<span data-ttu-id="2a82f-p112">Готовы [создать свою первую надстройку Word](word-add-ins.md)? Вы также можете воспользоваться нашим интерактивным [руководством по началу работы](/office/dev/add-ins/?product=Word). Используйте [манифест надстройки](../develop/add-in-manifests.md), чтобы указать ведущее приложение, имя, разрешения и другие сведения.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](/office/dev/add-ins/?product=Word). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="2a82f-155">Чтобы узнать больше о том, как создать качественную и привлекательную надстройку Word, см. [руководство по разработке](../design/add-in-design.md) и [рекомендации](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="2a82f-155">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="2a82f-156">После разработки надстройку можно [опубликовать](../publish/publish.md) в сетевой папке, каталоге приложений или AppSource.</span><span class="sxs-lookup"><span data-stu-id="2a82f-156">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="whats-coming-up-for-word-add-ins"></a><span data-ttu-id="2a82f-157">Над чем мы работаем?</span><span class="sxs-lookup"><span data-stu-id="2a82f-157">What's coming up for Word add-ins?</span></span>

<span data-ttu-id="2a82f-p113">Мы публикуем новые API для надстроек Word на странице [Открытые спецификации API](/office/dev/add-ins/reference/openspec), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями API JavaScript для Word мы работаем, и поделитесь своим мнением о проектируемых спецификациях.</span><span class="sxs-lookup"><span data-stu-id="2a82f-p113">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [API open specifications](/office/dev/add-ins/reference/openspec) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="2a82f-160">См. также</span><span class="sxs-lookup"><span data-stu-id="2a82f-160">See also</span></span>

* [<span data-ttu-id="2a82f-161">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2a82f-161">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="2a82f-162">Справочные материалы по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2a82f-162">Word JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
