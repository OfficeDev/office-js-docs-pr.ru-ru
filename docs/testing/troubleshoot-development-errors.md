---
title: Устранение ошибок разработки надстроек Office
description: Узнайте, как устранять ошибки разработки в надстройках Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409418"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="34f41-103">Устранение ошибок разработки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="34f41-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="34f41-104">Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки</span><span class="sxs-lookup"><span data-stu-id="34f41-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="34f41-105">Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="34f41-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="34f41-106">Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются</span><span class="sxs-lookup"><span data-stu-id="34f41-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="34f41-107">Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере.</span><span class="sxs-lookup"><span data-stu-id="34f41-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="34f41-108">Для Windows:</span><span class="sxs-lookup"><span data-stu-id="34f41-108">For Windows:</span></span>

<span data-ttu-id="34f41-109">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` и удалите содержимое папки `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , если она существует.</span><span class="sxs-lookup"><span data-stu-id="34f41-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="34f41-110">Для Mac</span><span class="sxs-lookup"><span data-stu-id="34f41-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="34f41-111">Для iOS</span><span class="sxs-lookup"><span data-stu-id="34f41-111">For iOS:</span></span>
<span data-ttu-id="34f41-p101">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="34f41-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="34f41-114">Изменения статических файлов, таких как JavaScript, HTML и CSS, не отображаются.</span><span class="sxs-lookup"><span data-stu-id="34f41-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="34f41-115">Браузер может кэшировать эти файлы.</span><span class="sxs-lookup"><span data-stu-id="34f41-115">The browser may be caching these files.</span></span> <span data-ttu-id="34f41-116">Чтобы избежать этого, отключите кэширование на стороне клиента при разработке.</span><span class="sxs-lookup"><span data-stu-id="34f41-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="34f41-117">Сведения будут зависеть от того, какой тип сервера вы используете.</span><span class="sxs-lookup"><span data-stu-id="34f41-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="34f41-118">В большинстве случаев необходимо добавить определенные заголовки в HTTP-ответы.</span><span class="sxs-lookup"><span data-stu-id="34f41-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="34f41-119">Мы предлагаем следующий набор заголовков:</span><span class="sxs-lookup"><span data-stu-id="34f41-119">We suggest the following set:</span></span>

- <span data-ttu-id="34f41-120">Cache-Control: "private, no-cache, no-store"</span><span class="sxs-lookup"><span data-stu-id="34f41-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="34f41-121">Pragma: "No-cache"</span><span class="sxs-lookup"><span data-stu-id="34f41-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="34f41-122">Expires: "-1"</span><span class="sxs-lookup"><span data-stu-id="34f41-122">Expires: "-1"</span></span>

<span data-ttu-id="34f41-123">Пример использования на сервере Node.JS Express см. в [этом файле app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="34f41-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="34f41-124">Пример использования в проекте ASP.NET см. в [этом файле cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="34f41-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="34f41-125">Если надстройка размещена на сервере Internet Information Server (IIS), можно также добавить указанные сведения в файл web.config.</span><span class="sxs-lookup"><span data-stu-id="34f41-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="34f41-126">Если сначала эти действия безуспешны, вам, возможно, потребуется очистить кэш браузера.</span><span class="sxs-lookup"><span data-stu-id="34f41-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="34f41-127">Сделайте это в интерфейсе браузера.</span><span class="sxs-lookup"><span data-stu-id="34f41-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="34f41-128">Иногда очистить кэш браузера Microsoft Edge, используя пользовательский интерфейс, не удается.</span><span class="sxs-lookup"><span data-stu-id="34f41-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="34f41-129">В таком случае выполните следующую команду в командной строке Windows.</span><span class="sxs-lookup"><span data-stu-id="34f41-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="34f41-130">Изменения, внесенные в значения свойств, не происходят и сообщение об ошибке не отображается</span><span class="sxs-lookup"><span data-stu-id="34f41-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="34f41-131">Ознакомьтесь с справочной документацией по свойству, чтобы проверить, доступно ли оно только для чтения.</span><span class="sxs-lookup"><span data-stu-id="34f41-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="34f41-132">Кроме того, [определения TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения.</span><span class="sxs-lookup"><span data-stu-id="34f41-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="34f41-133">Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="34f41-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="34f41-134">В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id). Просмотрите также, что [некоторые свойства не могут быть установлены напрямую](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span><span class="sxs-lookup"><span data-stu-id="34f41-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="34f41-135">Надстройка не работает на пограничной стороне, но работает в других браузерах</span><span class="sxs-lookup"><span data-stu-id="34f41-135">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="34f41-136">Ознакомьтесь с [разрешениями проблем Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span><span class="sxs-lookup"><span data-stu-id="34f41-136">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="34f41-137">Надстройка Excel вызывает ошибки, но не всегда</span><span class="sxs-lookup"><span data-stu-id="34f41-137">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="34f41-138">Возможные причины [: Устранение неполадок](../excel/excel-add-ins-troubleshooting.md) в надстройках Excel.</span><span class="sxs-lookup"><span data-stu-id="34f41-138">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="34f41-139">См. также</span><span class="sxs-lookup"><span data-stu-id="34f41-139">See also</span></span>

- [<span data-ttu-id="34f41-140">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="34f41-140">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="34f41-141">Загрузка неопубликованной надстройки Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="34f41-141">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="34f41-142">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="34f41-142">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="34f41-143">Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"</span><span class="sxs-lookup"><span data-stu-id="34f41-143">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="34f41-144">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="34f41-144">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="34f41-145">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="34f41-145">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="34f41-146">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="34f41-146">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
