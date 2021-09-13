---
title: Элемент PhoneSettings в файле манифеста
description: Элемент PhoneSettings указывает исходные параметры расположения и управления, которые применяются при применении почтовой надстройки на телефоне.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 1e52827a20ee95397541f7c1d54c732ff8f96ba5
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154547"
---
# <a name="phonesettings-element"></a>Элемент PhoneSettings

Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.

> [!IMPORTANT]
> Элемент доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и `PhoneSettings` Outlook 2013 Windows. Чтобы поддерживать Outlook Android и iOS, см. в приложении Надстройки [для Outlook Mobile.](../../outlook/outlook-mobile-addins.md)

**Тип надстройки:** почтовая

## <a name="syntax"></a>Синтаксис

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>Содержится в

[Form](form.md)

