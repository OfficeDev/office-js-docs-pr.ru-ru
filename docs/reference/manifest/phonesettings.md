---
title: Элемент PhoneSettings в файле манифеста
description: Элемент PhoneSettings указывает исходное расположение и параметры управления, которые применяются при использовании почтовой надстройки на телефоне.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 581a3ae71a58cd05aac52129a6f4395a60c20cef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720479"
---
# <a name="phonesettings-element"></a><span data-ttu-id="be191-103">Элемент PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="be191-103">PhoneSettings element</span></span>

<span data-ttu-id="be191-104">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.</span><span class="sxs-lookup"><span data-stu-id="be191-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="be191-105">`PhoneSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="be191-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="be191-106">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="be191-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="be191-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="be191-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="be191-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="be191-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="be191-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="be191-109">Contained in</span></span>

[<span data-ttu-id="be191-110">Form</span><span class="sxs-lookup"><span data-stu-id="be191-110">Form</span></span>](form.md)

