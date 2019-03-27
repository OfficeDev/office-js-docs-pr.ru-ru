---
title: Указание ведущих приложений Office и требований к API
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 9e1f92ae6708392157ccaa061cf0a064f3b5ece0
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870291"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="99a85-102">Указание ведущих приложений Office и требований к API</span><span class="sxs-lookup"><span data-stu-id="99a85-102">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="99a85-p101">Работа надстройки Office может зависеть от ведущего приложения Office, набора обязательных элементов, элемента или версии API. Например, надстройка может:</span><span class="sxs-lookup"><span data-stu-id="99a85-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="99a85-105">работать в одном (Word или Excel) или нескольких приложениях Office;</span><span class="sxs-lookup"><span data-stu-id="99a85-105">Run in a single Office application (Word or Excel), or several applications.</span></span>

- <span data-ttu-id="99a85-p102">использовать API JavaScript, доступные только в некоторых версиях Office. Например, можно создать надстройку Excel 2016 на базе новых API JavaScript для Excel;</span><span class="sxs-lookup"><span data-stu-id="99a85-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="99a85-108">работать только в версиях Office, которые поддерживают элементы API, используемые вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="99a85-108">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="99a85-109">Эта статья поможет вам разобраться, какие параметры следует выбрать для правильной работы надстройки и максимального охвата аудитории.</span><span class="sxs-lookup"><span data-stu-id="99a85-109">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="99a85-110">Общие сведения о том, на каких платформах поддерживаются надстройки Office, см. в [этой статье](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="99a85-110">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="99a85-111">В таблице ниже перечислены основные понятия, рассматриваемые в этой статье.</span><span class="sxs-lookup"><span data-stu-id="99a85-111">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="99a85-112">**Концепция**</span><span class="sxs-lookup"><span data-stu-id="99a85-112">**Concept**</span></span>|<span data-ttu-id="99a85-113">**Описание**</span><span class="sxs-lookup"><span data-stu-id="99a85-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="99a85-114">Приложение Office, ведущее приложение Office или ведущее приложение</span><span class="sxs-lookup"><span data-stu-id="99a85-114">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="99a85-p103">Приложение Office, используемое для запуска надстройки, например Word, Word Online, Excel и т. д.</span><span class="sxs-lookup"><span data-stu-id="99a85-p103">The Office application used to run your add-in. For example, Word, Word Online, Excel, and so on.</span></span>|
|<span data-ttu-id="99a85-117">Платформа</span><span class="sxs-lookup"><span data-stu-id="99a85-117">Platform</span></span>|<span data-ttu-id="99a85-118">Платформа ведущего приложения Office, например Office Online или Office для iPad.</span><span class="sxs-lookup"><span data-stu-id="99a85-118">Where the Office host runs, such as Office Online or Office for iPad.</span></span>|
|<span data-ttu-id="99a85-119">Набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="99a85-119">Requirement set</span></span>|<span data-ttu-id="99a85-p104">Именованная группа связанных элементов API. С помощью наборов обязательных элементов надстройка определяет, поддерживает ли ведущее приложение Office элементы API, которые она использует. Проще проверять поддержку набора обязательных элементов, а не отдельных элементов API. Поддержка набора обязательных элементов зависит от ведущего приложения Office и его версии. </span><span class="sxs-lookup"><span data-stu-id="99a85-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="99a85-124">Наборы обязательных элементов указываются в файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="99a85-124">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="99a85-125">Задавая наборы обязательных элементов в манифесте, вы указываете, какой минимальный уровень поддержки API должно обеспечить ведущее приложение Office, чтобы можно было запустить надстройку.</span><span class="sxs-lookup"><span data-stu-id="99a85-125">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in.</span></span> <span data-ttu-id="99a85-126">Надстройка не будет работать в ведущих приложениях Office, которые не поддерживают наборы обязательных элементов, указанные в манифесте, и не будет отображаться в разделе <span class="ui">Мои надстройки</span>. Это ограничивает доступность надстройки в коде с помощью проверок в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="99a85-126">Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.In code using runtime checks.</span></span> <span data-ttu-id="99a85-127">Полный список наборов требований см. в статье [Наборы обязательных элементов для надстроек Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="99a85-127">For the complete list of requirement sets, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>|
|<span data-ttu-id="99a85-128">Проверка в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="99a85-128">Runtime check</span></span>|<span data-ttu-id="99a85-p106">Проверка в среде выполнения, которая позволяет определить, поддерживает ли ведущее приложение Office наборы обязательных элементов или методы, которые использует надстройка. Чтобы запустить такую проверку, используйте оператор **if** с методом **isSetSupported**, наборами обязательных элементов или именами методов, которые не входят в набор обязательных элементов. Проверки в среде выполнения позволяют максимально расширить аудиторию надстройки. В отличие от наборов обязательных элементов, такие проверки не позволяют задать минимальный уровень поддержки API, который требуется для запуска надстройки. Вместо этого с помощью оператора **if** вы определяете, поддерживается ли элемент API, и если это так, добавляете в надстройку дополнительные функции. Если вы используете проверки в среде выполнения, ваша надстройка всегда будет отображаться в разделе **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="99a85-p106">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in. To perform a runtime check, you use an  **if** statement with the **isSetSupported** method, the requirement sets, or the method names that aren't part of a requirement set.Use runtime checks to ensure that your add-in reaches the broadest number of customers. Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run. Instead, you use the  **if** statement to determine whether an API member is supported. If it is, you can provide additional functionality in your add-in. Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="99a85-135">Перед началом работы</span><span class="sxs-lookup"><span data-stu-id="99a85-135">Before you begin</span></span>

<span data-ttu-id="99a85-p107">Надстройка должна использовать последнюю версию схемы манифеста надстройки. Если вы используете проверки в среде выполнения, используйте последнюю версию библиотеки API JavaScript для Office (office.js).</span><span class="sxs-lookup"><span data-stu-id="99a85-p107">Your add-in must use the most current version of the add-in manifest schema. If you use runtime checks in your add-in, ensure that you use the latest JavaScript API for Office (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="99a85-138">Выбор последней версии схема манифестов надстроек</span><span class="sxs-lookup"><span data-stu-id="99a85-138">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="99a85-p108">Ваша надстройка должна использовать схему манифеста 1.1. Настройте элемент **OfficeApp** в манифесте надстройки указанным ниже образом.</span><span class="sxs-lookup"><span data-stu-id="99a85-p108">Your add-in's manifest must use version 1.1 of the add-in manifest schema. Set the  **OfficeApp** element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a><span data-ttu-id="99a85-141">Выбор последней версии библиотеки API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="99a85-141">Specify the latest JavaScript API for Office library</span></span>

<span data-ttu-id="99a85-p109">Если вы используете проверки в среде выполнения, то вам необходимо ссылаться на последнюю версию библиотеки API JavaScript для Office из сети доставки содержимого. Для этого добавьте указанный ниже тег `script` в HTML-код. Чтобы всегда ссылаться на последнюю версию файла Office.js, используйте `/1/` в URL-адресе сети доставки содержимого.</span><span class="sxs-lookup"><span data-stu-id="99a85-p109">If you use runtime checks, reference the most current version of the JavaScript API for Office library from the content delivery network (CDN). To do this, add the following  `script` tag to your HTML. Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="99a85-145">Параметры для задания ведущих приложений Office или требований к API</span><span class="sxs-lookup"><span data-stu-id="99a85-145">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="99a85-p110">При указании ведущих приложений Office и требований к API необходимо учитывать несколько факторов. На следующей схеме показано, как выбрать правильный метод для надстройки.</span><span class="sxs-lookup"><span data-stu-id="99a85-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![Выбор самого подходящего варианта указания ведущих приложений Office или элементов API для надстройки](../images/options-for-office-hosts.png)

- <span data-ttu-id="99a85-p111">Если ваша надстройка работает в одном приложении Office, укажите элемент **Hosts** в манифесте. Дополнительные сведения см. в разделе [Задание элемента Hosts](#set-the-hosts-element).</span><span class="sxs-lookup"><span data-stu-id="99a85-p111">If your add-in runs in one Office host, set the **Hosts** element in the manifest. For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="99a85-p112">Чтобы задать минимальный набор обязательных элементов или элементы API, которые должно поддерживать ведущее приложение Office для запуска надстройки, задайте элемент **Requirements** в манифесте. Дополнительные сведения см. в разделе [Задание элемента Requirements в манифесте](#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="99a85-p112">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the  **Requirements** element in the manifest. For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="99a85-p113">Чтобы предоставить дополнительные функции, если определенные наборы обязательных элементов или элементы API доступны в ведущем приложении Office, выполните проверку в среде выполнения для кода JavaScript надстройки. Например, если надстройка выполняется в Excel 2016, используйте элементы нового API JavaScript для Excel, чтобы предоставить дополнительные функции. Дополнительные сведения см. в разделе [Использование проверок в среде выполнения для кода JavaScript](#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="99a85-p113">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code. For example, if your add-in runs in Excel 2016, use API members from the new JavaScript API for Excel to provide additional functionality. For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="99a85-156">Задание элемента Hosts</span><span class="sxs-lookup"><span data-stu-id="99a85-156">Set the Hosts element</span></span>

<span data-ttu-id="99a85-p114">Чтобы надстройка работала в одном ведущем приложении Office, используйте элементы **Hosts** и **Host** в манифесте. Если элемент **Hosts** не указан, надстройка будет работать во всех ведущих приложениях.</span><span class="sxs-lookup"><span data-stu-id="99a85-p114">To make your add-in run in one Office host application, use the  **Hosts** and **Host** elements in the manifest. If you don't specify the **Hosts** element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="99a85-159">Например, указанное ниже объявление **Hosts** и **Host** указывает, что надстройка будет работать с любым выпуском Excel, включая Excel для Windows, Excel Online и Excel для iPad.</span><span class="sxs-lookup"><span data-stu-id="99a85-159">For example, the following  **Hosts** and **Host** declaration specifies that the add-in will work with any release of Excel, which includes Excel for Windows, Excel Online, and Excel for iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="99a85-p115">Элемент **Hosts** может содержать один или несколько элементов **Host**. Элемент **Host** указывает ведущее приложение Office, в котором может работать ваша надстройка. Обязательному атрибуту **Name** можно присвоить одно из указанных ниже значений.</span><span class="sxs-lookup"><span data-stu-id="99a85-p115">The  **Hosts** element can contain one or more **Host** elements. The **Host** element specifies the Office host your add-in requires. The **Name** attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="99a85-163">Имя</span><span class="sxs-lookup"><span data-stu-id="99a85-163">Name</span></span>          | <span data-ttu-id="99a85-164">Ведущие приложения Office</span><span class="sxs-lookup"><span data-stu-id="99a85-164">Office host applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="99a85-165">Database</span><span class="sxs-lookup"><span data-stu-id="99a85-165">Database</span></span>      | <span data-ttu-id="99a85-166">Веб-приложения Access</span><span class="sxs-lookup"><span data-stu-id="99a85-166">Access web apps</span></span>                               |
| <span data-ttu-id="99a85-167">Document</span><span class="sxs-lookup"><span data-stu-id="99a85-167">Document</span></span>      | <span data-ttu-id="99a85-168">Word для Windows, Mac, iPad и Word Online</span><span class="sxs-lookup"><span data-stu-id="99a85-168">Word for Windows, Mac, iPad and Online</span></span>        |
| <span data-ttu-id="99a85-169">Mailbox</span><span class="sxs-lookup"><span data-stu-id="99a85-169">Mailbox</span></span>       | <span data-ttu-id="99a85-170">Outlook для Windows, Mac, Outlook в Интернете и Outlook.com</span><span class="sxs-lookup"><span data-stu-id="99a85-170">Outlook for Windows, Mac, Web and Outlook.com</span></span> |
| <span data-ttu-id="99a85-171">Presentation</span><span class="sxs-lookup"><span data-stu-id="99a85-171">Presentation</span></span>  | <span data-ttu-id="99a85-172">PowerPoint для Windows, Mac, iPad и PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="99a85-172">PowerPoint for Windows, Mac, iPad and Online</span></span>  |
| <span data-ttu-id="99a85-173">Project</span><span class="sxs-lookup"><span data-stu-id="99a85-173">Project</span></span>       | <span data-ttu-id="99a85-174">Project</span><span class="sxs-lookup"><span data-stu-id="99a85-174">Project</span></span>                                       |
| <span data-ttu-id="99a85-175">книга</span><span class="sxs-lookup"><span data-stu-id="99a85-175">Workbook</span></span>      | <span data-ttu-id="99a85-176">Excel для Windows, Mac, iPad и Excel Online</span><span class="sxs-lookup"><span data-stu-id="99a85-176">Excel Windows, Mac, iPad and Online</span></span>           |

> [!NOTE]
> <span data-ttu-id="99a85-p116">Атрибут `Name` указывает приложение Office, в котором может запускаться ваша надстройка. Приложения Office поддерживаются на разных платформах и работают на настольных ПК, в веб-браузерах, на планшетах и мобильных устройствах. Нельзя указать платформу, на которой можно запускать надстройку. Например, если вы укажете `Mailbox`, то для запуска надстройки можно будет использовать и Outlook, и Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="99a85-p116">The  `Name` attribute specifies the Office host application that can run your add-in. Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices. You can't specify which platform can be used to run your add-in. For example, if you specify `Mailbox`, both Outlook and Outlook Web App can be used to run your add-in.</span></span> 


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="99a85-181">Указание элемента Requirements в манифесте</span><span class="sxs-lookup"><span data-stu-id="99a85-181">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="99a85-p117">С помощью элемента **Requirements** можно задать минимальные наборы обязательных элементов или элементы API, которые должно поддерживать ведущее приложение Office для запуска надстройки. В элементе **Requirements** можно указать как наборы обязательных элементов, так и отдельные методы, используемые в надстройке. В версии 1.1 схемы манифестов надстроек элемент **Requirements** необязателен для всех надстроек, кроме надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="99a85-p117">The  **Requirements** element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in. The **Requirements** element can specify both requirement sets and individual methods used in your add-in. In version 1.1 of the add-in manifest schema, the **Requirements** element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="99a85-p118">Используйте элемент **Requirements**, только чтобы указать ключевые элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживают элементы API, указанные в элементе **Requirements**, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе **My Add-ins**. Рекомендуем сделать надстройку доступной на всех платформах ведущего приложения Office, например Excel для Windows, Excel Online и Excel для iPad. Чтобы надстройка была доступной во _всех_ приложениях Office и на всех платформах, используйте проверки в среде выполнения, а не элемент **Requirements**.</span><span class="sxs-lookup"><span data-stu-id="99a85-p118">Only use the **Requirements** element to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the **Requirements** element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the **Requirements** element.</span></span>

<span data-ttu-id="99a85-188">В примере кода ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих указанные ниже элементы.</span><span class="sxs-lookup"><span data-stu-id="99a85-188">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="99a85-189">Набор обязательных элементов **TableBindings** 1.1 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="99a85-189">**TableBindings** requirement set, which has a minimum version of 1.1.</span></span>

-  <span data-ttu-id="99a85-190">Набор обязательных элементов **OOXML** 1.1 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="99a85-190">**OOXML** requirement set, which has a minimum version of 1.1.</span></span>

-  <span data-ttu-id="99a85-191">Метод **Document.getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="99a85-191">**Document.getSelectedDataAsync** method.</span></span>

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- <span data-ttu-id="99a85-192">Элемент **Requirements** содержит дочерние элементы **Sets** и **Methods**.</span><span class="sxs-lookup"><span data-stu-id="99a85-192">The  **Requirements** element contains the **Sets** and **Methods** child elements.</span></span>

- <span data-ttu-id="99a85-p119">Элемент **Sets** может содержать один или несколько элементов **Set**. Параметр **DefaultMinVersion** задает значение **MinVersion** по умолчанию для всех дочерних элементов **Set**.</span><span class="sxs-lookup"><span data-stu-id="99a85-p119">The  **Sets** element can contain one or more **Set** elements. **DefaultMinVersion** specifies the default **MinVersion** value of all child **Set** elements.</span></span>

- <span data-ttu-id="99a85-195">Элемент **Set** указывает наборы обязательных элементов, которые ведущее приложение Office должно поддерживать для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="99a85-195">The  **Set** element specifies requirement sets that the Office host must support to run the add-in.</span></span> <span data-ttu-id="99a85-196">Атрибут **Name** указывает имя набора обязательных элементов.</span><span class="sxs-lookup"><span data-stu-id="99a85-196">The **Name** attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="99a85-197">Атрибут **MinVersion** указывает минимальную версию набора обязательных элементов.</span><span class="sxs-lookup"><span data-stu-id="99a85-197">The **MinVersion** specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="99a85-198">**MinVersion** переопределяет значение **DefaultMinVersion**.</span><span class="sxs-lookup"><span data-stu-id="99a85-198">**MinVersion** overrides the value of **DefaultMinVersion**.</span></span> <span data-ttu-id="99a85-199">Дополнительные сведения о наборах обязательных элементов и их версиях, к которым принадлежат элементы API, см. в статье [Наборы обязательных элементов для надстроек Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="99a85-199">For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

- <span data-ttu-id="99a85-p121">Элемент **Methods** может содержать один или несколько элементов **Method**. Элемент **Methods** не следует использовать с надстройками Outlook.</span><span class="sxs-lookup"><span data-stu-id="99a85-p121">The  **Methods** element can contain one or more **Method** elements. You can't use the **Methods** element with Outlook add-ins.</span></span>

- <span data-ttu-id="99a85-p122">Элемент **Method** задает отдельный метод, который должно поддерживать ведущее приложение Office, в котором работает надстройка. Атрибут **Name** обязателен и указывает имя метода с его родительским объектом.</span><span class="sxs-lookup"><span data-stu-id="99a85-p122">The  **Method** element specifies an individual method that must be supported in the Office host where your add-in runs. The **Name** attribute is required and specifies the name of the method qualified with its parent object.</span></span>


## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="99a85-204">Использование проверок в среде выполнения для кода JavaScript</span><span class="sxs-lookup"><span data-stu-id="99a85-204">Use runtime checks in your JavaScript code</span></span>


<span data-ttu-id="99a85-p123">Если ведущее приложение Office поддерживает определенные наборы обязательных элементов, вы можете добавить в надстройку дополнительные функции. Например, если надстройка работает в Word 2016, вы можете использовать в ней новые API JavaScript для Word. Для этого используйте метод **isSetSupported** с именем набора обязательных элементов. В среде выполнения метод **isSetSupported** определяет, поддерживается ли набор обязательных элементов ведущим приложением Office, в котором запущена надстройка. Если он поддерживается, то метод **isSetSupported** возвращает значение **true** и запускает дополнительный код, который использует элементы API из этого набора. Если приложение Office не поддерживает набор обязательных элементов, метод **isSetSupported** возвращает значение **false**, и дополнительный код не запускается. В коде ниже показан синтаксис, который необходимо использовать с методом **isSetSupported**.</span><span class="sxs-lookup"><span data-stu-id="99a85-p123">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host. For example, you might want to use the new Word JavaScript APIs Word in your existing add-in if your add-in runs in Word 2016. To do this, you use the  **isSetSupported** method with the name of the requirement set. **isSetSupported** determines, at runtime, whether the Office host running the add-in supports the requirement set. If the requirement set is supported, **isSetSupported** returns **true** and runs the additional code that uses the API members from that requirement set. If the Office host doesn't support the requirement set, **isSetSupported** returns **false** and the additional code won't run. The following code shows the syntax to use with **isSetSupported**.</span></span>


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```

-  <span data-ttu-id="99a85-212">_RequirementSetName_ (обязательный параметр) — это строка, представляющая имя набора обязательных элементов.</span><span class="sxs-lookup"><span data-stu-id="99a85-212">_RequirementSetName_ (required) is a string that represents the name of the requirement set.</span></span> <span data-ttu-id="99a85-213">Дополнительные сведения о доступных наборах обязательных элементов см. в статье [Наборы обязательных элементов для надстроек Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="99a85-213">For more information about available requirement sets, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>
    
-  <span data-ttu-id="99a85-214">_VersionNumber_ (необязательный параметр) — это версия набора обязательных элементов.</span><span class="sxs-lookup"><span data-stu-id="99a85-214">_VersionNumber_ (optional) is the version of the requirement set.</span></span>
    
<span data-ttu-id="99a85-p125">В Excel 2016 или Word 2016 для наборов обязательных элементов **ExcelAPI** или **WordAPI** используйте метод **isSetSupported**. Метод **isSetSupported**, а также наборы обязательных элементов **ExcelAPI** и **WordAPI** доступны в последней версии файла Office.js в CDN. Если вы не используете файл Office.js из CDN, надстройка может создавать исключения, так как метод **isSetSupported** не будет определен. Дополнительные сведения см. в статье [Выбор последней версии библиотеки API JavaScript для Office](#specify-the-latest-javascript-api-for-office-library).</span><span class="sxs-lookup"><span data-stu-id="99a85-p125">In Excel 2016 or Word 2016, use  **isSetSupported** with the **ExcelAPI** or **WordAPI** requirement sets. The **isSetSupported** method, and the **ExcelAPI** and **WordAPI** requirement sets, are available in the latest Office.js file available from the CDN. If you don't use Office.js from the CDN, your add-in might generate exceptions because **isSetSupported** will be undefined. For more information, see [Specify the latest JavaScript API for Office library](#specify-the-latest-javascript-api-for-office-library).</span></span> 


> [!NOTE]
> <span data-ttu-id="99a85-p126">Метод **isSetSupported** не работает в Outlook и Outlook Web App. Чтобы использовать проверку в среде выполнения в Outlook или Outlook Web App, используйте способ, описанный в разделе [Проверки в среде выполнения с использованием методов, не входящих в набор обязательных элементов](#runtime-checks-using-methods-not-in-a-requirement-set).</span><span class="sxs-lookup"><span data-stu-id="99a85-p126">**isSetSupported** does not work in Outlook or Outlook Web App. To use a runtime check in Outlook or Outlook Web App, use the technique described in [Runtime checks using methods not in a requirement set](#runtime-checks-using-methods-not-in-a-requirement-set).</span></span>

<span data-ttu-id="99a85-221">В приведенном ниже примере кода показано, как функциональность надстройки может отличаться в ведущих приложениях Office, поддерживающих разные наборы обязательных элементов или элементы API.</span><span class="sxs-lookup"><span data-stu-id="99a85-221">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>

```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="99a85-222">Проверки в среде выполнения с использованием методов, не входящих в набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="99a85-222">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="99a85-223">Некоторые элементы API не входят в наборы обязательных элементов.</span><span class="sxs-lookup"><span data-stu-id="99a85-223">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="99a85-224">Это относится только к тем элементам API, которые входят в пространства имен [JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office) (все элементы в Office.), и не относится к элементам API, принадлежащим к пространствам имен API JavaScript для Word (все элементы в Word.) или [API JavaScript для надстроек Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) (все элементы в Excel.).</span><span class="sxs-lookup"><span data-stu-id="99a85-224">This only applies to API members that are part of the [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) namespace (anything under Office.), not API members that belong to the Word JavaScript API (anything in Word.) or [Excel add-ins JavaScript API reference](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) (anything in Excel.) namespaces.</span></span> <span data-ttu-id="99a85-225">Если надстройка зависит от метода, не входящего в набор обязательных элементов, вы можете использовать проверку в среде выполнения, чтобы определить, поддерживается ли метод ведущим приложением Office, как показано в примере кода ниже.</span><span class="sxs-lookup"><span data-stu-id="99a85-225">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example.</span></span> <span data-ttu-id="99a85-226">Список всех методов, не входящих в набор обязательных элементов, см. в статье [Наборы обязательных элементов для надстроек Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="99a85-226">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

> [!NOTE]
> <span data-ttu-id="99a85-227">Рекомендуем ограничить использование этого типа проверки в среде выполнения в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="99a85-227">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="99a85-228">В примере кода ниже показано, как проверить, поддерживает ли ведущее приложение метод **document.setSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="99a85-228">The following code example checks whether the host supports  **document.setSelectedDataAsync**.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="99a85-229">См. также</span><span class="sxs-lookup"><span data-stu-id="99a85-229">See also</span></span>

- [<span data-ttu-id="99a85-230">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="99a85-230">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="99a85-231">Наборы обязательных элементов для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="99a85-231">Office Add-in requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="99a85-232">Word-Add-in-Get-Set-EditOpen-XML </span><span class="sxs-lookup"><span data-stu-id="99a85-232">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
