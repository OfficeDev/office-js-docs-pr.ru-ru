---
title: Элемент Hosts в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 606073977366e37ecc4419f468f01bfb25647a7d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452027"
---
# <a name="hosts-element"></a><span data-ttu-id="ff1fa-102">Элемент Hosts</span><span class="sxs-lookup"><span data-stu-id="ff1fa-102">Hosts element</span></span>

<span data-ttu-id="ff1fa-p101">Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров.</span><span class="sxs-lookup"><span data-stu-id="ff1fa-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="ff1fa-105">При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="ff1fa-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="ff1fa-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ff1fa-106">Child elements</span></span>

|  <span data-ttu-id="ff1fa-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="ff1fa-107">Element</span></span> |  <span data-ttu-id="ff1fa-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ff1fa-108">Required</span></span>  |  <span data-ttu-id="ff1fa-109">Описание</span><span class="sxs-lookup"><span data-stu-id="ff1fa-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ff1fa-110">Host</span><span class="sxs-lookup"><span data-stu-id="ff1fa-110">Host</span></span>](host.md)    |  <span data-ttu-id="ff1fa-111">Да</span><span class="sxs-lookup"><span data-stu-id="ff1fa-111">Yes</span></span>   |  <span data-ttu-id="ff1fa-112">Описывает ведущее приложение и его параметры.</span><span class="sxs-lookup"><span data-stu-id="ff1fa-112">Describes a host and its settings.</span></span> |
