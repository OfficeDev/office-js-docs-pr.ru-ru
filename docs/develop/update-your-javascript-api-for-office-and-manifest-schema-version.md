---
title: Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1
description: Обновление до версии 1.1 файлов JavaScript (Office.js и JS-файлов приложения) и файла проверки манифеста надстройки в проекте надстройки Office.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 6acd08a388b162cec4ac30fdfb256adc980d9e69
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626756"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="3a5e3-103">Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="3a5e3-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="3a5e3-104">В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="3a5e3-105">Проекты, созданные в Visual Studio 2019, уже используют версию 1,1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-105">Projects created in Visual Studio 2019 will already use version 1.1.</span></span> <span data-ttu-id="3a5e3-106">Однако для версии 1.1 периодически выпускаются незначительные обновления, которые можно применить с помощью методов, описанных в этой статье.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="3a5e3-107">Использование последних версий файлов в проекте</span><span class="sxs-lookup"><span data-stu-id="3a5e3-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="3a5e3-108">Если вы используете Visual Studio для разработки надстройки, для использования новейших элементов API JavaScript для Office и [компонентов версии 1.1 манифеста надстройки](../develop/add-in-manifests.md) (которая проверяется по сравнению с offappmanifest-1.1. xsd) необходимо скачать Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-108">If you use Visual Studio to develop your add-in, to use the newest API members of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2019.</span></span> <span data-ttu-id="3a5e3-109">Чтобы скачать Visual Studio 2019, посетите [страницу Visual Studio IDE](https://visualstudio.microsoft.com/vs/).</span><span class="sxs-lookup"><span data-stu-id="3a5e3-109">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="3a5e3-110">Во время установки потребуется выбрать рабочую нагрузку разработки Office и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="3a5e3-111">Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="3a5e3-112">Чтобы запустить надстройку, разработанную с использованием новых и обновленных компонентов манифеста надстройки и интерфейса API Office.js, ваши клиенты должны использовать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также при необходимости SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанными серверными продуктами, Пакет обновления 1 (SP1) для Exchange Server 2013 или аналогичные размещенные в сети продукты: Office 365, SharePoint Online и Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="3a5e3-113">Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="3a5e3-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="3a5e3-114">Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем</span><span class="sxs-lookup"><span data-stu-id="3a5e3-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="3a5e3-115">Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов</span><span class="sxs-lookup"><span data-stu-id="3a5e3-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="3a5e3-116">Описание пакета обновления 1 для Exchange Server 2013</span><span class="sxs-lookup"><span data-stu-id="3a5e3-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="3a5e3-117">Обновление проекта надстройки Office, созданного в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="3a5e3-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="3a5e3-118">Для проектов, созданных до выпуска версии 1.1 библиотеки JavaScript API для Office и схемы манифеста надстройки, вы можете обновить файлы проекта, используя **диспетчер пакетов NuGet**, а затем добавить ссылки на них в HTML-страницы надстройки.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-118">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="3a5e3-119">Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="3a5e3-120">Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии</span><span class="sxs-lookup"><span data-stu-id="3a5e3-120">Update the JavaScript API for Office library files in your project to the newest release</span></span>
<span data-ttu-id="3a5e3-121">Приведенные ниже действия приведут к обновлению файлов библиотеки Office. js до последней версии.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-121">The following steps will update your Office.js library files to the latest version.</span></span> <span data-ttu-id="3a5e3-122">В этом разделе описано, как использовать Visual Studio 2019, но они аналогичны предыдущим версиям Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-122">The steps use Visual Studio 2019, but they are similar for previous versions of Visual Studio.</span></span>

1. <span data-ttu-id="3a5e3-123">В Visual Studio 2019 откройте или создайте новый проект **надстройки Office** .</span><span class="sxs-lookup"><span data-stu-id="3a5e3-123">In Visual Studio 2019, open or create a new  **Office Add-in** project.</span></span>
2. <span data-ttu-id="3a5e3-124">Выберите **Средства** > **Диспетчер пакетов NuGet** > **Управление пакетами Nuget для решения**.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-124">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="3a5e3-125">Выберите вкладку **Обновления**.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-125">Choose the **Updates** tab.</span></span>
4. <span data-ttu-id="3a5e3-126">Выберите Microsoft.Office.js.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-126">Select Microsoft.Office.js.</span></span> <span data-ttu-id="3a5e3-127">Убедитесь, что источник пакета находится в **NuGet.org**.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-127">Ensure the package source is from **nuget.org**.</span></span>
5. <span data-ttu-id="3a5e3-128">В левой области выберите **установить** и завершить процесс обновления пакета.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-128">In the left pane, choose **Install** and complete the package update process.</span></span>

<span data-ttu-id="3a5e3-129">Вам потребуется выполнить несколько дополнительных действий, чтобы завершить обновление.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="3a5e3-130">В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="3a5e3-131">`/1/` в `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="3a5e3-132">Обновление схемы манифеста в проекте до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="3a5e3-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="3a5e3-133">В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="3a5e3-134">После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы   **Capabilities** и **Capability** и заменить их либо элементами [Hosts](/office/dev/add-ins/reference/manifest/hosts) и [Host](/office/dev/add-ins/reference/manifest/host), либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="3a5e3-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="3a5e3-135">Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE</span><span class="sxs-lookup"><span data-stu-id="3a5e3-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="3a5e3-136">Если вы создали проект до выпуска схемы манифеста надстройки и API JavaScript для Office версии 1.1, обновите HTML-страницы вашей надстройки, чтобы они ссылались на CDN библиотеки версии 1.1, а также обновите файл манифеста надстройки, чтобы использовалась схема версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-136">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="3a5e3-137">Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="3a5e3-138">Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и JS-файлов для конкретной надстройки), чтобы разрабатывать надстройку Office (ссылки на CDN для Office.js позволяют скачивать необходимые файлы во время выполнения). Если вам нужны файлы библиотеки, то вы можете скачать их с помощью [служебной программы командной строки NuGet](https://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js`.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-138">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="3a5e3-139">Чтобы получить копию XSD (определения схемы XML) для манифеста надстройки версии 1.1, см. статью [Справочник по схеме манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="3a5e3-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="3a5e3-140">Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии</span><span class="sxs-lookup"><span data-stu-id="3a5e3-140">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="3a5e3-141">Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="3a5e3-142">В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="3a5e3-143">`/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний добавочный выпуск Office.js версии 1.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="3a5e3-144">Обновление схемы манифеста в проекте до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="3a5e3-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="3a5e3-145">В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.</span><span class="sxs-lookup"><span data-stu-id="3a5e3-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="3a5e3-146">После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы   **Capabilities** и **Capability** и заменить их либо элементами [Hosts](/office/dev/add-ins/reference/manifest/hosts) и [Host](/office/dev/add-ins/reference/manifest/host), либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="3a5e3-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](/office/dev/add-ins/reference/manifest/hosts) and [Host](/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3a5e3-147">См. также</span><span class="sxs-lookup"><span data-stu-id="3a5e3-147">See also</span></span>

- <span data-ttu-id="3a5e3-148">[Указание ведущих приложений Office и требований к API](specify-office-hosts-and-api-requirements.md) ]</span><span class="sxs-lookup"><span data-stu-id="3a5e3-148">[Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="3a5e3-149">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="3a5e3-149">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="3a5e3-150">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="3a5e3-150">JavaScript API for Office</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="3a5e3-151">Справка по схеме для манифестов надстроек Office (версия 1.1)</span><span class="sxs-lookup"><span data-stu-id="3a5e3-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
