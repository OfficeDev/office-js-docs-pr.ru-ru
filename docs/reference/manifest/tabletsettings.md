---
title: Элемент TabletSettingst в файле манифеста
description: Элемент TabletSettings указывает параметры элементов управления, которые применяются при использовании почтовой надстройки на планшете.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 71b7aed6b2906a8695ac1c13e93ba60da1aa56ec
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215084"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="b2f35-103">Элемент TabletSettingst</span><span class="sxs-lookup"><span data-stu-id="b2f35-103">TabletSettings element</span></span>

<span data-ttu-id="b2f35-104">Задает параметры управления, которые применяются при использовании вашей почтовой надстройки на планшете.</span><span class="sxs-lookup"><span data-stu-id="b2f35-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b2f35-105">`TabletSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="b2f35-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="b2f35-106">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="b2f35-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="b2f35-107">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="b2f35-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b2f35-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b2f35-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b2f35-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b2f35-109">Contained in</span></span>

[<span data-ttu-id="b2f35-110">Form</span><span class="sxs-lookup"><span data-stu-id="b2f35-110">Form</span></span>](form.md)
