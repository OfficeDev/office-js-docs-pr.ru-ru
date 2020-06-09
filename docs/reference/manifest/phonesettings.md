---
title: Элемент PhoneSettings в файле манифеста
description: Элемент PhoneSettings указывает исходное расположение и параметры управления, которые применяются при использовании почтовой надстройки на телефоне.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: d7957e23a77a0f837366e5cedc0e0f350b5635c8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611486"
---
# <a name="phonesettings-element"></a><span data-ttu-id="587e6-103">Элемент PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="587e6-103">PhoneSettings element</span></span>

<span data-ttu-id="587e6-104">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.</span><span class="sxs-lookup"><span data-stu-id="587e6-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="587e6-105">`PhoneSettings`Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="587e6-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="587e6-106">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="587e6-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="587e6-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="587e6-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="587e6-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="587e6-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="587e6-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="587e6-109">Contained in</span></span>

[<span data-ttu-id="587e6-110">Form</span><span class="sxs-lookup"><span data-stu-id="587e6-110">Form</span></span>](form.md)

