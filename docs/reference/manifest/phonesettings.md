---
title: Элемент PhoneSettings в файле манифеста
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 4614c86af865e5242657f47e21e6786545a616b6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165540"
---
# <a name="phonesettings-element"></a><span data-ttu-id="17d8a-102">Элемент PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="17d8a-102">PhoneSettings element</span></span>

<span data-ttu-id="17d8a-103">Задает исходное расположение и параметры элемента управления, которые применяются при использовании почтовой надстройки на телефоне.</span><span class="sxs-lookup"><span data-stu-id="17d8a-103">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="17d8a-104">`PhoneSettings` Элемент доступен только в классическом приложении Outlook в Интернете (как правило, подключенный к старым версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="17d8a-104">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="17d8a-105">Для поддержки Outlook в Android и iOS, ознакомьтесь со статьей надстройки [для Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="17d8a-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="17d8a-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="17d8a-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="17d8a-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="17d8a-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="17d8a-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="17d8a-108">Contained in</span></span>

[<span data-ttu-id="17d8a-109">Form</span><span class="sxs-lookup"><span data-stu-id="17d8a-109">Form</span></span>](form.md)

