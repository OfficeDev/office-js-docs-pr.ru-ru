---
title: Элемент PhoneSettings в файле манифеста
description: Элемент PhoneSettings указывает исходное расположение и параметры управления, которые применяются при использовании почтовой надстройки на телефоне.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: cfdf8efc262c1a75142eb06dd71383579598a86f
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215049"
---
# <a name="phonesettings-element"></a>Элемент PhoneSettings

Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.

> [!IMPORTANT]
> `PhoneSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows. Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).

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

