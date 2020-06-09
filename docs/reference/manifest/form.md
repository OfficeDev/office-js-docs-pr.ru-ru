---
title: Элемент Form в файле манифеста
description: Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: c9cd1d9104fc51edc84149ef677c4308dfb1a9f5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611857"
---
# <a name="form-element"></a><span data-ttu-id="4818b-103">Элемент Form</span><span class="sxs-lookup"><span data-stu-id="4818b-103">Form element</span></span>

<span data-ttu-id="4818b-104">Параметры взаимодействия с пользователем для форм, которые почтовая надстройка будет использовать при работе на определенном устройства (настольном компьютере, планшете или телефоне).</span><span class="sxs-lookup"><span data-stu-id="4818b-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4818b-105">`DesktopSettings`Элементы, `TabletSettings` и, Кроме того, `PhoneSettings` доступны только в классическом приложении Outlook в Интернете (как правило, подключаются к предыдущим версиям локального сервера Exchange Server) и Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="4818b-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="4818b-106">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="4818b-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4818b-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4818b-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="4818b-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="4818b-108">Contained in</span></span>

[<span data-ttu-id="4818b-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="4818b-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="4818b-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="4818b-110">Can contain</span></span>

|<span data-ttu-id="4818b-111">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="4818b-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="4818b-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="4818b-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="4818b-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="4818b-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="4818b-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="4818b-114">PhoneSettings</span></span>](phonesettings.md)|
