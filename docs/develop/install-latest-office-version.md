---
title: Установка последней версии Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925236"
---
# <a name="install-the-latest-version-of-office-2016"></a>Установка последней версии Office 2016

Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Как получать последние сборки раньше других

Чтобы получать последние сборки Office 2016 раньше других: 

- Если вы подписаны на Office 365 для дома, Office 365 персональный или Office 365 для студентов, [примите участие в программе предварительной оценки Office](https://products.office.com/office-insider).
- Если вы пользуетесь Office 365 для бизнеса, прочитайте статью [Установка сборки раннего выпуска для клиентов Office 365 для бизнеса](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Если вы используете Office 2016 для Mac:
    - Запустите программу Office 2016 для Mac.
    - Выберите пункт **Проверить наличие обновлений** в меню "Справка".
    - В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office. 

## <a name="get-the-latest-build"></a>Как получить последнюю сборку

Чтобы получить последнюю сборку Office 2016: 

1. Скачайте [средство развертывания Office 2016](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.
3. Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > Команды может выполняться долго, при этом ход ее выполнения нигде не отображается.

По завершении процесса установки у вас будут последние версии приложений Office 2016. Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл**  >  **Учетная запись**. В разделе "Обновления Office" над номером версии должна быть надпись Office Insiders.

![Снимок экрана, на котором показаны сведения о продукте с надписью "Участники программы предварительной оценки Office"](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office

Сведения о минимальных сборках продуктов для каждой платформы см. в следующих статьях:

- [Наборы обязательных элементов API JavaScript для Word](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для OneNote](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [Наборы обязательных элементов API диалоговых окон](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [Наборы обязательных элементов общего API для Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
