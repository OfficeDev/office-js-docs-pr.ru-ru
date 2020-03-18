---
title: Элемент TabletSettingst в файле манифеста
description: Элемент TabletSettings указывает параметры элементов управления, которые применяются при использовании почтовой надстройки на планшете.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 2b8b372d27274d89d3aed4b5bacb9faa4893fda5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717861"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="ede24-103">Элемент TabletSettingst</span><span class="sxs-lookup"><span data-stu-id="ede24-103">TabletSettings element</span></span>

<span data-ttu-id="ede24-104">Задает параметры управления, которые применяются при использовании вашей почтовой надстройки на планшете.</span><span class="sxs-lookup"><span data-stu-id="ede24-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ede24-105">`TabletSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="ede24-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="ede24-106">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="ede24-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="ede24-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="ede24-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ede24-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ede24-108">Syntax</span></span>

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a><span data-ttu-id="ede24-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="ede24-109">Contained in</span></span>

[<span data-ttu-id="ede24-110">Form</span><span class="sxs-lookup"><span data-stu-id="ede24-110">Form</span></span>](form.md)

