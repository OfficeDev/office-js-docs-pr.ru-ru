---
title: Обзор надстроек Word
description: Изучите основы надстроек Word
ms.date: 11/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: b27de57c821fe8278ac8a6e097aeb8e5be0cdb7b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717329"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="7065b-103">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="7065b-103">Word add-ins overview</span></span>

<span data-ttu-id="7065b-p101">Хотите создать решение для автоматического составления документов или привязки и доступа к данным в документе Word из других источников? Чтобы расширить возможности клиентов Word на компьютере с Windows, Mac или в облаке, используйте платформу надстроек Office, которая включает API JavaScript для Word и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="7065b-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="7065b-p102">На [платформе надстроек Office](../overview/office-add-ins.md) можно разрабатывать не только надстройки Word. Используя команды надстроек, вы можете расширять интерфейс Word и запускать области задач, которые выполняют сценарий JavaScript, взаимодействующий с содержимым документа. Любой код, который работает в браузере, будет работать в надстройке Word. Надстройки, взаимодействующие с содержимым документа Word, создают запросы на совершение действий с объектами Word и синхронизацию состояния этих объектов.</span><span class="sxs-lookup"><span data-stu-id="7065b-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="7065b-p103">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource, она должна соответствовать [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и [статье о доступности надстроек Office в ведущих приложениях](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="7065b-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="7065b-114">Ниже показан пример надстройки Word, работающей в области задач.</span><span class="sxs-lookup"><span data-stu-id="7065b-114">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="7065b-115">*Рис. 1. Надстройка, работающая в области задач Word*</span><span class="sxs-lookup"><span data-stu-id="7065b-115">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Надстройка, работающая в области задач Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="7065b-p104">Надстройка Word может (1) отправлять запросы в документ Word и (2) обновлять, удалять или перемещать абзац, используя JavaScript для доступа к объекту paragraph. Например, в приведенном ниже коде показано, как добавить в абзац новое предложение.</span><span class="sxs-lookup"><span data-stu-id="7065b-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

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

<span data-ttu-id="7065b-p105">Для размещения надстройки Word можно использовать любой веб-сервер, в частности ASP.NET, NodeJS и Python. Используйте любимую клиентскую платформу — Ember, Backbone, Angular, React —для разработки своего решения; или продолжайте работу с VanillaJS. Для [аутентификации](../develop/overview-authn-authz.md) и размещения приложения можно использовать Azure.</span><span class="sxs-lookup"><span data-stu-id="7065b-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.</span></span>

<span data-ttu-id="7065b-p106">API JavaScript для Word предоставляют приложению доступ к объектам и метаданным документа Word. С помощью этих API можно создавать надстройки, предназначенные для:</span><span class="sxs-lookup"><span data-stu-id="7065b-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="7065b-123">Word 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="7065b-123">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="7065b-124">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="7065b-124">Word on the web</span></span>
* <span data-ttu-id="7065b-125">Word 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="7065b-125">Word 2016 or later on Mac</span></span>
* <span data-ttu-id="7065b-126">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="7065b-126">Word on iPad</span></span>

<span data-ttu-id="7065b-p107">Написанные вами надстройки будут работать во всех версиях Word на различных платформах. Дополнительные сведения см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="7065b-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="7065b-129">API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="7065b-129">JavaScript APIs for Word</span></span>

<span data-ttu-id="7065b-130">Для взаимодействия с объектами и метаданными в документе Word можно использовать два набора API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7065b-130">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="7065b-131">Первый — [общий API](/javascript/api/office), представленный в Office 2013.</span><span class="sxs-lookup"><span data-stu-id="7065b-131">The first is the [Common API](/javascript/api/office), which was introduced in Office 2013.</span></span> <span data-ttu-id="7065b-132">Многие объекты общего API можно использовать в надстройках, размещенных в двух или более клиентах Office.</span><span class="sxs-lookup"><span data-stu-id="7065b-132">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="7065b-133">В этом API широко используются обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="7065b-133">This API uses callbacks extensively.</span></span>

<span data-ttu-id="7065b-p109">Второй — [API JavaScript для Word](/javascript/api/word). Это строго типизированная объектная модель, с помощью которой можно создавать надстройки Word, предназначенные для Word 2016 для Mac и Windows. Эта объектная модель использует обещания и предоставляет доступ к объектам Word, в частности [Body](/javascript/api/word/word.body), [ContentControl](/javascript/api/word/word.contentcontrol), [InlinePicture](/javascript/api/word/word.inlinepicture) и [Paragraph](/javascript/api/word/word.paragraph). API JavaScript для Word включает определения TypeScript и файлы vsdoc, чтобы вы могли получать подсказки кода в своей интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="7065b-p109">The second is the [Word JavaScript API](/javascript/api/word). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="7065b-p110">В настоящее время все клиенты Word поддерживают общий API JavaScript для Office, а большинство из них поддерживают и API JavaScript для Word. Дополнительные сведения о поддерживаемых клиентах см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="7065b-p110">Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API. For details about supported clients, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

<span data-ttu-id="7065b-p111">Рекомендуем начать с API JavaScript для Word, так как с объектной моделью проще работать. Используйте API JavaScript для Word, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="7065b-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="7065b-142">получить доступ к объектам в документе Word.</span><span class="sxs-lookup"><span data-stu-id="7065b-142">Access the objects in a Word document.</span></span>

<span data-ttu-id="7065b-143">Используйте общий API JavaScript для Office, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="7065b-143">Use the shared Office JavaScript API when you need to:</span></span>

* <span data-ttu-id="7065b-144">создать надстройки для Word 2013;</span><span class="sxs-lookup"><span data-stu-id="7065b-144">Target Word 2013.</span></span>
* <span data-ttu-id="7065b-145">выполнить начальные действия для приложения;</span><span class="sxs-lookup"><span data-stu-id="7065b-145">Perform initial actions for the application.</span></span>
* <span data-ttu-id="7065b-146">проверить поддерживаемый набор требований;</span><span class="sxs-lookup"><span data-stu-id="7065b-146">Check the supported requirement set.</span></span>
* <span data-ttu-id="7065b-147">получить доступ к метаданным документа, его параметрам и сведениям о среде;</span><span class="sxs-lookup"><span data-stu-id="7065b-147">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="7065b-148">создать привязку к разделам документа и записать события;</span><span class="sxs-lookup"><span data-stu-id="7065b-148">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="7065b-149">использовать пользовательские XML-части;</span><span class="sxs-lookup"><span data-stu-id="7065b-149">Use custom XML parts.</span></span>
* <span data-ttu-id="7065b-150">открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7065b-150">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7065b-151">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7065b-151">Next steps</span></span>

<span data-ttu-id="7065b-p112">Готовы [создать свою первую надстройку Word](word-add-ins.md)? Используйте [манифест надстройки](../develop/add-in-manifests.md), чтобы указать ведущее приложение, имя, разрешения и другие сведения.</span><span class="sxs-lookup"><span data-stu-id="7065b-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="7065b-155">Чтобы узнать больше о том, как создать качественную и привлекательную надстройку Word, см. [руководство по разработке](../design/add-in-design.md) и [рекомендации](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="7065b-155">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="7065b-156">После разработки надстройку можно [опубликовать](../publish/publish.md) в сетевой папке, каталоге приложений или AppSource.</span><span class="sxs-lookup"><span data-stu-id="7065b-156">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="7065b-157">См. также</span><span class="sxs-lookup"><span data-stu-id="7065b-157">See also</span></span>

* [<span data-ttu-id="7065b-158">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7065b-158">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="7065b-159">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7065b-159">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="7065b-160">Справочные материалы по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="7065b-160">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)
