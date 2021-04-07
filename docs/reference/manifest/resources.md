---
title: Элемент Resources в файле манифеста
description: В элементе Resources находятся значки, строки и URL-адреса для узла VersionOverrides.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: bdf73420345ca4d054438bfba5217254e6682e6d
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604619"
---
# <a name="resources-element"></a>Элемент Resources

Содержит значки, строки и URL-адреса для узла [VersionOverrides](versionoverrides.md). Элемент манифеста указывает ресурс с помощью атрибута **id**. Это позволяет сократить размер манифеста, особенно когда имеются версии ресурсов для разных языковых стандартов. Атрибут **id** должен быть уникальным в пределах манифеста и не может быть длиннее 32 символов.

Каждый ресурс может иметь один или несколько дочерних элементов **Override**, позволяющих указать другой ресурс для определенного языкового стандарта.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Тип  |  Описание  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  Предоставляет URL-адрес HTTPS изображения значка. |
|  **Urls**                |  url     |  Предоставляет URL-адрес HTTPS расположения. URL-адрес не может быть длиннее 2048 символов. |
|  **ShortStrings** |  string  |  Текст для элементов **Label** и **Title**. Каждая **строка** содержит не более 125 символов.|
|  **LongStrings**  |  string  | Текст для атрибутов **Description**. Каждая **строка** содержит не более 250 символов.|

> [!NOTE]
> Для всех URL-адресов в элементах **Image** и **Url** необходимо использовать протокол SSL.

### <a name="images"></a>изображения;

Каждый значок должен иметь три элемента **Изображения,** по одному для каждого из трех обязательных размеров:

- 16 x 16
- 32x32
- 80x80

Кроме того, поддерживаются (но не требуются) указанные ниже дополнительные размеры.

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT]
>
> - Если это изображение является представителем значка надстройки, см. в приложении [Create effective listings in AppSource и Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) для размера и других требований.
> - Для оптимальной работы Outlook требуется кэшировать ресурсы изображений. Поэтому сервер, на котором размещен ресурс изображения, не должен добавлять директивы CACHE-CONTROL в заголовок ответа. Это приведет к тому, что Outlook автоматически заменит универсальное или стандартное изображение.

## <a name="resources-examples"></a>Примеры ресурсов

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
