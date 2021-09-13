---
title: Элемент TabletSettingst в файле манифеста
description: Элемент TabletSettings указывает параметры управления, которые применяются при добавлении почты на планшете.
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3d7ace7fe9258ee32f3f5507d35b35ae026ef5eb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151215"
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
