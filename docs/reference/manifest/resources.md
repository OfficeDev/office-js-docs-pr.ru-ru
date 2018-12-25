---
title: Элемент Resources в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0707df137d075a9922836e5d960216d089c56675
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433903"
---
# <a name="resources-element"></a><span data-ttu-id="28ca7-102">Элемент Resources</span><span class="sxs-lookup"><span data-stu-id="28ca7-102">Resources element</span></span>

<span data-ttu-id="28ca7-p101">Содержит значки, строки и URL-адреса для узла [VersionOverrides](versionoverrides.md). Элемент манифеста указывает ресурс с помощью атрибута **id**. Это позволяет сократить размер манифеста, особенно когда имеются версии ресурсов для разных языковых стандартов. Атрибут **id** должен быть уникальным в пределах манифеста и не может быть длиннее 32 символов.</span><span class="sxs-lookup"><span data-stu-id="28ca7-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="28ca7-107">Каждый ресурс может иметь один или несколько дочерних элементов **Override**, позволяющих указать другой ресурс для определенного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="28ca7-107">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="28ca7-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="28ca7-108">Child elements</span></span>

|  <span data-ttu-id="28ca7-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="28ca7-109">Element</span></span> |  <span data-ttu-id="28ca7-110">Тип</span><span class="sxs-lookup"><span data-stu-id="28ca7-110">Type</span></span>  |  <span data-ttu-id="28ca7-111">Описание</span><span class="sxs-lookup"><span data-stu-id="28ca7-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="28ca7-112">Images</span><span class="sxs-lookup"><span data-stu-id="28ca7-112">Images</span></span>](#images)            |  <span data-ttu-id="28ca7-113">image</span><span class="sxs-lookup"><span data-stu-id="28ca7-113">image</span></span>   |  <span data-ttu-id="28ca7-114">Предоставляет URL-адрес HTTPS изображения значка.</span><span class="sxs-lookup"><span data-stu-id="28ca7-114">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="28ca7-115">**Urls**</span><span class="sxs-lookup"><span data-stu-id="28ca7-115">**Urls**</span></span>                |  <span data-ttu-id="28ca7-116">url</span><span class="sxs-lookup"><span data-stu-id="28ca7-116">url</span></span>     |  <span data-ttu-id="28ca7-p102">Предоставляет URL-адрес HTTPS расположения. URL-адрес не может быть длиннее 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="28ca7-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="28ca7-119">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="28ca7-119">**ShortStrings**</span></span> |  <span data-ttu-id="28ca7-120">string</span><span class="sxs-lookup"><span data-stu-id="28ca7-120">string</span></span>  |  <span data-ttu-id="28ca7-p103">Текст для элементов **Label** и **Title**. Каждая **строка** содержит не более 125 символов.</span><span class="sxs-lookup"><span data-stu-id="28ca7-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="28ca7-123">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="28ca7-123">**LongStrings**</span></span>  |  <span data-ttu-id="28ca7-124">string</span><span class="sxs-lookup"><span data-stu-id="28ca7-124">string</span></span>  | <span data-ttu-id="28ca7-p104">Текст для атрибутов **Description**. Каждая**строка** содержит не более 250 символов.</span><span class="sxs-lookup"><span data-stu-id="28ca7-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="28ca7-127">Для всех URL-адресов в элементах **Image** и **Url** необходимо использовать протокол SSL.</span><span class="sxs-lookup"><span data-stu-id="28ca7-127">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="28ca7-128">Images</span><span class="sxs-lookup"><span data-stu-id="28ca7-128">Images</span></span>
<span data-ttu-id="28ca7-129">У каждого значка должно быть три элемента **Images**, по одному на каждый из трех обязательных размеров:</span><span class="sxs-lookup"><span data-stu-id="28ca7-129">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="28ca7-130">16 x 16</span><span class="sxs-lookup"><span data-stu-id="28ca7-130">16x16</span></span>
- <span data-ttu-id="28ca7-131">32x32</span><span class="sxs-lookup"><span data-stu-id="28ca7-131">32x32</span></span>
- <span data-ttu-id="28ca7-132">80x80</span><span class="sxs-lookup"><span data-stu-id="28ca7-132">80x80</span></span>

<span data-ttu-id="28ca7-133">Кроме того, поддерживаются (но не требуются) указанные ниже дополнительные размеры.</span><span class="sxs-lookup"><span data-stu-id="28ca7-133">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="28ca7-134">20x20</span><span class="sxs-lookup"><span data-stu-id="28ca7-134">20x20</span></span>
- <span data-ttu-id="28ca7-135">24x24</span><span class="sxs-lookup"><span data-stu-id="28ca7-135">24x24</span></span>
- <span data-ttu-id="28ca7-136">40x40</span><span class="sxs-lookup"><span data-stu-id="28ca7-136">40x40</span></span>
- <span data-ttu-id="28ca7-137">48x48</span><span class="sxs-lookup"><span data-stu-id="28ca7-137">48x48</span></span>
- <span data-ttu-id="28ca7-138">64x64</span><span class="sxs-lookup"><span data-stu-id="28ca7-138">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="28ca7-139">Для оптимальной работы Outlook требуется кэшировать ресурсы изображений.</span><span class="sxs-lookup"><span data-stu-id="28ca7-139">Important:  Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="28ca7-140">Поэтому сервер, на котором размещен ресурс изображения, не должен добавлять директивы CACHE-CONTROL в заголовок ответа.</span><span class="sxs-lookup"><span data-stu-id="28ca7-140">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="28ca7-141">Это приведет к тому, что Outlook автоматически заменит универсальное или стандартное изображение.</span><span class="sxs-lookup"><span data-stu-id="28ca7-141">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="28ca7-142">Примеры ресурсов</span><span class="sxs-lookup"><span data-stu-id="28ca7-142">Resources examples</span></span> 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
