---
title: Установка последней версии Office
description: Сведения о том, как получать последние сборки Office раньше других.
ms.date: 12/04/2017
ms.openlocfilehash: 0e6e147144757004575fa086e1066b7cdf133ee8
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505792"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="2ada2-103">Установка последней версии Office</span><span class="sxs-lookup"><span data-stu-id="2ada2-103">Install the latest version of Office</span></span>

<span data-ttu-id="2ada2-104">Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других.</span><span class="sxs-lookup"><span data-stu-id="2ada2-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="2ada2-105">Как получать последние сборки раньше других</span><span class="sxs-lookup"><span data-stu-id="2ada2-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="2ada2-106">Чтобы получать последние сборки Office раньше других:</span><span class="sxs-lookup"><span data-stu-id="2ada2-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="2ada2-107">Если вы подписаны на Office 365 для дома, Office 365 персональный или Office 365 для студентов, [примите участие в программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="2ada2-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="2ada2-108">Если вы пользуетесь Office 365 для бизнеса, прочитайте статью [Установка сборки раннего выпуска для клиентов Office 365 для бизнеса](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span><span class="sxs-lookup"><span data-stu-id="2ada2-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="2ada2-109">Если вы используете Office для Mac:</span><span class="sxs-lookup"><span data-stu-id="2ada2-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="2ada2-110">Запустите программу Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="2ada2-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="2ada2-111">Выберите пункт **Проверить наличие обновлений** в меню "Справка".</span><span class="sxs-lookup"><span data-stu-id="2ada2-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="2ada2-112">В окне "Автоматическое обновление (Майкрософт)" установите флажок для присоединения к программе предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="2ada2-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="2ada2-113">Как получить последнюю сборку</span><span class="sxs-lookup"><span data-stu-id="2ada2-113">Get the latest build</span></span>

<span data-ttu-id="2ada2-114">Чтобы получить последнюю сборку Office:</span><span class="sxs-lookup"><span data-stu-id="2ada2-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="2ada2-115">Скачайте [средство развертывания Office](https://www.microsoft.com/download/details.aspx?id=49117).</span><span class="sxs-lookup"><span data-stu-id="2ada2-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="2ada2-p101">Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.</span><span class="sxs-lookup"><span data-stu-id="2ada2-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="2ada2-118">Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span><span class="sxs-lookup"><span data-stu-id="2ada2-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="2ada2-119">Выполните следующую команду от имени администратора:  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="2ada2-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="2ada2-120">Команда может выполняться долго, при этом ход ее выполнения нигде не отображается.</span><span class="sxs-lookup"><span data-stu-id="2ada2-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="2ada2-121">По завершении процесса установки будут установлены последние приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2ada2-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="2ada2-122">Чтобы убедиться в том, что у вас установлена последняя сборка, выберите **Файл** > **Учетная запись** из любого приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2ada2-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="2ada2-123">В разделе "Обновления Office" над номером версии должна быть надпись Office Insiders.</span><span class="sxs-lookup"><span data-stu-id="2ada2-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Снимок экрана, на котором показаны сведения о продукте с надписью "Участники программы предварительной оценки Office"](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="2ada2-125">Минимальные сборки Office, которые могут использовать наборы требований API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2ada2-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="2ada2-126">Сведения о минимальных сборках продуктов для каждой платформы для наборов требований API см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="2ada2-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="2ada2-127">Наборы требований API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="2ada2-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="2ada2-128">Наборы требований API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2ada2-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="2ada2-129">Наборы требований API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="2ada2-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="2ada2-130">Наборы требований Dialog API</span><span class="sxs-lookup"><span data-stu-id="2ada2-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="2ada2-131">Стандартные наборы требований API для Office</span><span class="sxs-lookup"><span data-stu-id="2ada2-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
