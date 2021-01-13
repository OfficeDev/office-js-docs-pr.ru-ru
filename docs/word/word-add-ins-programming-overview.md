---
title: Обзор надстроек Word
description: Изучите основы надстроек Word.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 2dba67cabaf11d6e10560ba3dbe5babde3ed0c92
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840008"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="e3bb1-103">Обзор надстроек Word</span><span class="sxs-lookup"><span data-stu-id="e3bb1-103">Word add-ins overview</span></span>

<span data-ttu-id="e3bb1-p101">Хотите создать решение для автоматического составления документов или привязки и доступа к данным в документе Word из других источников? Чтобы расширить возможности клиентов Word на компьютере с Windows, Mac или в облаке, используйте платформу надстроек Office, которая включает API JavaScript для Word и API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="e3bb1-p102">На [платформе надстроек Office](../overview/office-add-ins.md) можно разрабатывать не только надстройки Word. Используя команды надстроек, вы можете расширять интерфейс Word и запускать области задач, которые выполняют сценарий JavaScript, взаимодействующий с содержимым документа. Любой код, который работает в браузере, будет работать в надстройке Word. Надстройки, взаимодействующие с содержимым документа Word, создают запросы на совершение действий с объектами Word и синхронизацию состояния этих объектов.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="e3bb1-112">Ниже показан пример надстройки Word, работающей в области задач.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-112">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="e3bb1-113">*Рис. 1. Надстройка, работающая в области задач Word*</span><span class="sxs-lookup"><span data-stu-id="e3bb1-113">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Надстройка, работающая в области задач Word](../images/word-add-in-show-host-client.png)

<span data-ttu-id="e3bb1-p103">Надстройка Word может (1) отправлять запросы в документ Word и (2) обновлять, удалять или перемещать абзац, используя JavaScript для доступа к объекту paragraph. Например, в приведенном ниже коде показано, как добавить в абзац новое предложение.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p103">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

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

<span data-ttu-id="e3bb1-p104">Для размещения надстройки Word можно использовать любой веб-сервер, в частности ASP.NET, NodeJS и Python. Используйте любимую клиентскую платформу — Ember, Backbone, Angular, React —для разработки своего решения; или продолжайте работу с VanillaJS. Для [аутентификации](../develop/overview-authn-authz.md) и размещения приложения можно использовать Azure.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p104">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.</span></span>

<span data-ttu-id="e3bb1-p105">API JavaScript для Word предоставляют приложению доступ к объектам и метаданным документа Word. С помощью этих API можно создавать надстройки, предназначенные для:</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p105">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="e3bb1-121">Word 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="e3bb1-121">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="e3bb1-122">Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="e3bb1-122">Word on the web</span></span>
* <span data-ttu-id="e3bb1-123">Word 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="e3bb1-123">Word 2016 or later on Mac</span></span>
* <span data-ttu-id="e3bb1-124">Word для iPad</span><span class="sxs-lookup"><span data-stu-id="e3bb1-124">Word on iPad</span></span>

<span data-ttu-id="e3bb1-p106">Написанные вами надстройки будут работать во всех версиях Word на различных платформах. Дополнительные сведения см. в статье [Доступность клиентских приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p106">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="e3bb1-127">API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3bb1-127">JavaScript APIs for Word</span></span>

<span data-ttu-id="e3bb1-128">Для взаимодействия с объектами и метаданными в документе Word можно использовать два набора API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-128">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="e3bb1-129">Первый — [общий API](/javascript/api/office), представленный в Office 2013.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-129">The first is the [Common API](/javascript/api/office), which was introduced in Office 2013.</span></span> <span data-ttu-id="e3bb1-130">Многие объекты общего API можно использовать в надстройках, размещенных в двух или более клиентах Office.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-130">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="e3bb1-131">В этом API широко используются обратные вызовы.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-131">This API uses callbacks extensively.</span></span>

<span data-ttu-id="e3bb1-p108">Второй — [API JavaScript для Word](/javascript/api/word). Это [зависящая от приложений модель API](../develop/application-specific-api-model.md), которая появилась в Word 2016. Это строго типизированная объектная модель, с помощью которой можно создавать надстройки Word, предназначенные для Word 2016 для Mac и Windows. Эта объектная модель использует обещания и предоставляет доступ к объектам Word, в частности [Body](/javascript/api/word/word.body), [ContentControl](/javascript/api/word/word.contentcontrol), [InlinePicture](/javascript/api/word/word.inlinepicture) и [Paragraph](/javascript/api/word/word.paragraph). API JavaScript для Word включает определения TypeScript и файлы vsdoc, чтобы вы могли получать подсказки кода в своей интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p108">The second is the [Word JavaScript API](/javascript/api/word). This is a [application-specific API model](../develop/application-specific-api-model.md) that was introduced with Word 2016. It's a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="e3bb1-137">В настоящее время все клиенты Word поддерживают общий API JavaScript для Office, а большинство из них поддерживают и API JavaScript для Word.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-137">Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API.</span></span> <span data-ttu-id="e3bb1-138">Дополнительные сведения о поддерживаемых клиентах см. в статье [Доступность клиентских приложений и платформ Office для надстроек Office](../overview/office-add-in-availability.md)</span><span class="sxs-lookup"><span data-stu-id="e3bb1-138">For details about supported clients, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

<span data-ttu-id="e3bb1-p110">Рекомендуем начать с API JavaScript для Word, так как с объектной моделью проще работать. Используйте API JavaScript для Word, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p110">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="e3bb1-141">получить доступ к объектам в документе Word.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-141">Access the objects in a Word document.</span></span>

<span data-ttu-id="e3bb1-142">Используйте общий API JavaScript для Office, если вам нужно:</span><span class="sxs-lookup"><span data-stu-id="e3bb1-142">Use the shared Office JavaScript API when you need to:</span></span>

* <span data-ttu-id="e3bb1-143">создать надстройки для Word 2013;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-143">Target Word 2013.</span></span>
* <span data-ttu-id="e3bb1-144">выполнить начальные действия для приложения;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-144">Perform initial actions for the application.</span></span>
* <span data-ttu-id="e3bb1-145">проверить поддерживаемый набор требований;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-145">Check the supported requirement set.</span></span>
* <span data-ttu-id="e3bb1-146">получить доступ к метаданным документа, его параметрам и сведениям о среде;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-146">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="e3bb1-147">создать привязку к разделам документа и записать события;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-147">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="e3bb1-148">использовать пользовательские XML-части;</span><span class="sxs-lookup"><span data-stu-id="e3bb1-148">Use custom XML parts.</span></span>
* <span data-ttu-id="e3bb1-149">открыть диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-149">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e3bb1-150">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="e3bb1-150">Next steps</span></span>

<span data-ttu-id="e3bb1-p111">Готовы [создать свою первую надстройку Word](../quickstarts/word-quickstart.md)? Используйте [манифест надстройки](../develop/add-in-manifests.md), чтобы указать ведущее приложение, имя, разрешения и другие сведения.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-p111">Ready to create your first Word add-in? See [Build your first Word add-in](../quickstarts/word-quickstart.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="e3bb1-154">Чтобы узнать больше о том, как создать качественную и привлекательную надстройку Word, см. [руководство по разработке](../design/add-in-design.md) и [рекомендации](../concepts/add-in-development-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="e3bb1-154">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="e3bb1-155">После разработки надстройку можно [опубликовать](../publish/publish.md) в сетевой папке, каталоге приложений или AppSource.</span><span class="sxs-lookup"><span data-stu-id="e3bb1-155">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="e3bb1-156">См. также</span><span class="sxs-lookup"><span data-stu-id="e3bb1-156">See also</span></span>

* [<span data-ttu-id="e3bb1-157">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e3bb1-157">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="e3bb1-158">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="e3bb1-158">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
* [<span data-ttu-id="e3bb1-159">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e3bb1-159">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="e3bb1-160">Справочные материалы по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="e3bb1-160">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)