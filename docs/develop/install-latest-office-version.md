---
title: Установка последней версии Office
description: Сведения о том, как получать последние сборки Office раньше других.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: adfed2e5e35e2ad86295faafc2ffed91cf728bcd
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128327"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="cf09f-103">Установка последней версии Office</span><span class="sxs-lookup"><span data-stu-id="cf09f-103">Install the latest version of Office</span></span>

<span data-ttu-id="cf09f-104">Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других.</span><span class="sxs-lookup"><span data-stu-id="cf09f-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="cf09f-105">Как получать последние сборки раньше других</span><span class="sxs-lookup"><span data-stu-id="cf09f-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="cf09f-106">Чтобы получать последние сборки Office раньше других:</span><span class="sxs-lookup"><span data-stu-id="cf09f-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="cf09f-107">Если вы подписаны на Office 365 для дома, Office 365 персональный или Office 365 для студентов, [примите участие в программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="cf09f-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="cf09f-108">Если вы пользуетесь Office 365 для бизнеса, прочитайте статью [Установка сборки раннего выпуска для клиентов Office 365 для бизнеса](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="cf09f-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="cf09f-109">Если вы используете Office для Mac:</span><span class="sxs-lookup"><span data-stu-id="cf09f-109">If you're running Office on a Mac:</span></span>
    - <span data-ttu-id="cf09f-110">Запустите приложение Office.</span><span class="sxs-lookup"><span data-stu-id="cf09f-110">Start the Office application.</span></span>
    - <span data-ttu-id="cf09f-111">Выберите пункт **Проверить наличие обновлений** в меню "Справка".</span><span class="sxs-lookup"><span data-stu-id="cf09f-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="cf09f-112">В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="cf09f-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="cf09f-113">Как получить последнюю сборку</span><span class="sxs-lookup"><span data-stu-id="cf09f-113">Get the latest build</span></span>

<span data-ttu-id="cf09f-114">Чтобы получить последнюю сборку Office:</span><span class="sxs-lookup"><span data-stu-id="cf09f-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="cf09f-115">Скачайте [средство развертывания Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="cf09f-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="cf09f-p101">Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="cf09f-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="cf09f-118">Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="cf09f-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="cf09f-119">Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="cf09f-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

    > [!NOTE]
    > <span data-ttu-id="cf09f-120">Команда может выполняться долго, при этом ход ее выполнения нигде не отображается.</span><span class="sxs-lookup"><span data-stu-id="cf09f-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="cf09f-121">По завершении процесса установки у вас будут последние версии приложений Office.</span><span class="sxs-lookup"><span data-stu-id="cf09f-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="cf09f-122">Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл** > **Учетная запись**.</span><span class="sxs-lookup"><span data-stu-id="cf09f-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="cf09f-123">В разделе "Обновления Office" над номером версии должна быть надпись "Предварительная оценка Office".</span><span class="sxs-lookup"><span data-stu-id="cf09f-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Снимок экрана, на котором показаны сведения о продукте с надписью "Предварительная оценка Office"](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="cf09f-125">Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="cf09f-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="cf09f-126">Сведения о минимальных сборках продуктов для каждой платформы см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="cf09f-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="cf09f-127">Наборы обязательных элементов API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="cf09f-127">Word JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="cf09f-128">Наборы обязательных элементов API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="cf09f-128">Excel JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="cf09f-129">Наборы обязательных элементов API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="cf09f-129">OneNote JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="cf09f-130">Наборы обязательных элементов API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="cf09f-130">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="cf09f-131">Наборы обязательных элементов общего API для Office</span><span class="sxs-lookup"><span data-stu-id="cf09f-131">Office Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
