---
title: Элемент TabletSettingst в файле манифеста
description: Элемент TabletSettings указывает параметры управления, которые применяются при добавлении почты на планшете.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937552"
---
# <a name="tabletsettings-element"></a>Элемент TabletSettingst

Задает параметры управления, которые применяются при использовании вашей почтовой надстройки на планшете.

> [!IMPORTANT]
> Элемент доступен только в классических Outlook в Интернете (обычно подключенных к старым версиям локального Exchange сервера) и `TabletSettings` Outlook 2013 Windows. Чтобы поддерживать Outlook Android и iOS, см. в приложении Надстройки [для Outlook Mobile.](../../outlook/outlook-mobile-addins.md)

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
