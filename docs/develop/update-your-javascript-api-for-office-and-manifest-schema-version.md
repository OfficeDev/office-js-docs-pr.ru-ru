---
title: Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1
description: Обновление файлов JavaScript (Office.js и JS-файлы приложения) и файла проверки манифеста надстройки в вашем проекте надстройки Office до версии 1.1.
ms.date: 12/04/2017
ms.openlocfilehash: e58239a4e67871eb955d7fc205e26d0eb95af327
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004940"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="df86b-103">Обновление библиотеки API JavaScript для Office до последней версии и схемы манифеста надстройки до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="df86b-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="df86b-104">В этой статье рассказывается, как обновить файлы JavaScript (Office.js и JS-файлы для конкретной надстройки) и файл проверки манифеста надстройки в проекте надстройки Office до версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="df86b-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="df86b-105">Использование последних версий файлов в проекте</span><span class="sxs-lookup"><span data-stu-id="df86b-105">Use the most up-to-date project files</span></span>

<span data-ttu-id="df86b-106">Если для разработки надстройки вы используете Visual Studio, то чтобы можно было применять [самые новые элементы API](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js) в API JavaScript для Office и [возможности манифеста надстройки версии 1.1](../develop/add-in-manifests.md) (который проверяется на соответствие offappmanifest-1.1.xsd), вам потребуется скачать и установить [Visual Studio 2015 и последнюю версию Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).</span><span class="sxs-lookup"><span data-stu-id="df86b-106">If you use Visual Studio to develop your add-in, to use the [newest API members](https://docs.microsoft.com/javascript/office/what's-changed-in-the-javascript-api-for-office?view=office-js) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download and install the [Visual Studio 2015 and the latest Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).</span></span>

<span data-ttu-id="df86b-107">Если вы используете текстовый редактор или другую интегрированную среду разработки, отличную от Visual Studio, чтобы разработать надстройка, обновите ссылки на CDN для файла Office.js и версию схемы, на которую ссылается манифест приложения для Office.</span><span class="sxs-lookup"><span data-stu-id="df86b-107">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="df86b-108">Чтобы запустить надстройку, разработанную с использованием новых и обновленных компонентов манифеста надстройки и интерфейса API Office.js, ваши клиенты должны использовать локальные продукты Office 2013 с пакетом обновления 1 (SP1) или более поздней версии, а также при необходимости SharePoint Server 2013 с пакетом обновления 1 (SP1) и связанными серверными продуктами, Пакет обновления 1 (SP1) для Exchange Server 2013 или аналогичные размещенные в сети продукты: Office 365, SharePoint Online и Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="df86b-108">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="df86b-109">Сведения о том, как скачать Office, SharePoint и Exchange с пакетом обновления 1, см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="df86b-109">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="df86b-110">Список всех пакетов обновления 1 (SP1) для Microsoft Office 2013 и связанных продуктов для настольных систем</span><span class="sxs-lookup"><span data-stu-id="df86b-110">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](http://support.microsoft.com/kb/2850036)
    
- [<span data-ttu-id="df86b-111">Список всех пакетов обновления 1 (SP1) для Microsoft SharePoint Server 2013 и связанных серверных продуктов</span><span class="sxs-lookup"><span data-stu-id="df86b-111">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](http://support.microsoft.com/kb/2850035)
    
- [<span data-ttu-id="df86b-112">Описание пакета обновления 1 для Exchange Server 2013</span><span class="sxs-lookup"><span data-stu-id="df86b-112">Description of Exchange Server 2013 Service Pack 1</span></span>](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="df86b-113">Обновление проекта надстройки Office, созданного в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="df86b-113">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="df86b-114">Для проектов, созданных до выпуска версии 1.1 библиотеки JavaScript API для Office и схемы манифеста надстройки, вы можете обновить файлы проекта, используя **диспетчер пакетов NuGet**, а затем добавить ссылки на них в HTML-страницы надстройки.</span><span class="sxs-lookup"><span data-stu-id="df86b-114">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="df86b-115">Обратите внимание, что процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать Office.js и схему манифеста надстройки версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="df86b-115">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="df86b-116">Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии</span><span class="sxs-lookup"><span data-stu-id="df86b-116">Update the JavaScript API for Office library files in your project to the newest release</span></span>


1. <span data-ttu-id="df86b-117">В Visual Studio 2015 откройте или создайте проект **Надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="df86b-117">In Visual Studio 2015, open or create a new  **Office Add-in** project.</span></span>
    
      - <span data-ttu-id="df86b-118">В расположенной слева области щелкните **Обновить** и завершите процесс обновления пакета.</span><span class="sxs-lookup"><span data-stu-id="df86b-118">In the left pane, choose **Update** and complete the package update process.</span></span>
    
      - <span data-ttu-id="df86b-119">Перейдите к этапу 6.</span><span class="sxs-lookup"><span data-stu-id="df86b-119">Go to step 6.</span></span>
    
2. <span data-ttu-id="df86b-120">Выберите **Средства**  >  **Диспетчер пакетов NuGet**  >  **Управление пакетами Nuget для решения**.</span><span class="sxs-lookup"><span data-stu-id="df86b-120">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
    
3. <span data-ttu-id="df86b-p101">В **диспетчере пакетов NuGet** выберите **nuget.org** в качестве **источника пакетов** и **Доступны обновления** в поле **Фильтр**. Затем выберите файл Microsoft.Office.js.</span><span class="sxs-lookup"><span data-stu-id="df86b-p101">In the  **NuGet Package Manager**, select  **nuget.org** for **Package source** and **Upgrade available** for **Filter**. and select Microsoft.Office.js.</span></span>
    
4. <span data-ttu-id="df86b-123">В области слева выберите **Обновить** и завершите обновление пакета.</span><span class="sxs-lookup"><span data-stu-id="df86b-123">In the left pane, choose **Update** and complete the package update process.</span></span>
    
5. <span data-ttu-id="df86b-124">В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="df86b-124">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > <span data-ttu-id="df86b-125">Цифра `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний накопительный выпуск Office.js версии 1.</span><span class="sxs-lookup"><span data-stu-id="df86b-125">NOTE The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>   


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="df86b-126">Обновление схемы манифеста в проекте до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="df86b-126">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="df86b-127">В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.</span><span class="sxs-lookup"><span data-stu-id="df86b-127">In your Add-in's Manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="df86b-128">После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) и [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="df86b-128">NOTE After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) and [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="df86b-129">Обновление проекта надстройки Office, созданного с помощью текстового редактора или другой среды IDE</span><span class="sxs-lookup"><span data-stu-id="df86b-129">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="df86b-130">Если вы создали проект до выпуска схемы манифеста надстройки и API JavaScript для Office версии 1.1, обновите HTML-страницы вашей надстройки, чтобы они ссылались на CDN библиотеки версии 1.1, а также обновите файл манифеста надстройки, чтобы использовалась схема версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="df86b-130">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="df86b-131">Процесс обновления применяется к _проектам по отдельности_. Вам потребуется повторить его для каждого проекта надстройки, в котором вы хотите использовать файл Office.js и схему манифеста надстройки версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="df86b-131">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="df86b-132">Вам не нужны локальные копии файлов API JavaScript для Office (Office.js и JS-файлов для конкретной надстройки), чтобы разрабатывать надстройку Office (ссылки на CDN для Office.js позволяют скачивать необходимые файлы во время выполнения). Если вам нужны файлы библиотеки, то вы можете скачать их с помощью [служебной программы командной строки NuGet](http://docs.nuget.org/consume/installing-nuget) и `Install-Package Microsoft.Office.js`.</span><span class="sxs-lookup"><span data-stu-id="df86b-132">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](http://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE] 
> <span data-ttu-id="df86b-133">Чтобы получить копию файла XSD (определение схемы XML) для манифеста надстройки версии 1.1, см. запись в статье [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="df86b-133">NOTE To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="df86b-134">Обновление файлов библиотеки API JavaScript для Office в проекте до последней версии</span><span class="sxs-lookup"><span data-stu-id="df86b-134">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="df86b-135">Откройте HTML-страницы надстройки в текстовом редакторе или интегрированной среде разработки.</span><span class="sxs-lookup"><span data-stu-id="df86b-135">Open the HTML pages for your add-in in your text editor or IDE.</span></span>
    
2. <span data-ttu-id="df86b-136">В теге **head** HTML-страниц надстройки закомментируйте или удалите все ссылки на скрипт office.js и добавьте ссылки на обновленную библиотеку API JavaScript для Office, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="df86b-136">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > <span data-ttu-id="df86b-137">Цифра `/1/` перед `office.js` в URL-адресе CDN указывает на то, что необходимо использовать последний накопительный выпуск Office.js версии 1.</span><span class="sxs-lookup"><span data-stu-id="df86b-137">NOTE The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="df86b-138">Обновление схемы манифеста в проекте до версии 1.1</span><span class="sxs-lookup"><span data-stu-id="df86b-138">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="df86b-139">В файле манифеста надстройки обновите атрибут **xmlns** элемента **OfficeApp**, заменив значение версии на `1.1` и оставив все атрибуты, кроме **xmlns**, без изменений.</span><span class="sxs-lookup"><span data-stu-id="df86b-139">In your Add-in's Manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="df86b-140">После обновления схемы манифеста надстройки до версии 1.1 вам потребуется удалить элементы **Capabilities** и **Capability** и заменить их элементами [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) и [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) либо [элементами Requirements и Requirement](specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="df86b-140">NOTE After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/javascript/office/manifest/hosts?view=office-js) and [Host](https://docs.microsoft.com/javascript/office/manifest/host?view=office-js) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>
    

## <a name="see-also"></a><span data-ttu-id="df86b-141">См. также</span><span class="sxs-lookup"><span data-stu-id="df86b-141">See also</span></span>

- [<span data-ttu-id="df86b-142">Указание ведущих приложений Office и элементов API</span><span class="sxs-lookup"><span data-stu-id="df86b-142">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md) 
- [<span data-ttu-id="df86b-143">Общие сведения об интерфейсе API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="df86b-143">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="df86b-144">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="df86b-144">JavaScript API for Office</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)   
- [<span data-ttu-id="df86b-145">Справка по схеме для манифестов надстроек Office (версия 1.1)</span><span class="sxs-lookup"><span data-stu-id="df86b-145">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
    
