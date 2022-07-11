---
title: Загрузка неопубликованных надстроек Office на Mac для тестирования
description: Протестируйте надстройку Office на Компьютере Mac, выполнив загрузку неопубликованных приложений.
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38ed5f5dba2d379b6137a098240021bd642d6e11
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713231"
---
# <a name="sideload-office-add-ins-on-mac-for-testing"></a>Загрузка неопубликованных надстроек Office на Mac для тестирования

Чтобы узнать, как надстройка будет работать в Office на Mac, можно загрузить манифест надстройки неопубликованным образом. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.

> [!NOTE]
> Соответствующие действия касательно надстройки Outlook приведены в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-mac"></a>Предварительные требования (Office для Mac)

- Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).

- Word для Mac версии 15.18 (160109).

- Excel для Mac версии 15.19 (160206).

- PowerPoint для Mac версии 15.24 (160614).

- XML-файл манифеста для надстройки, которую вы хотите протестировать.

## <a name="sideload-an-add-in-in-office-on-mac"></a>Загрузка неопубликованной надстройки в Office для Mac

1. Используйте **Finder** для загрузки неопубликованного файла манифеста. Откройте **Finder и** введите COMMAND+SHIFT+G, чтобы открыть диалоговое окно "Перейти **к папке** ".

1. Введите один из следующих путей к файлам в зависимости от приложения, которое вы хотите использовать для загрузки неопубликованных приложений. Если папки `wef` нет на компьютере, создайте ее.

    - Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > В остальных шагах описывается загрузка неопубликоваемой надстройки Word.

1. Скопируйте файл манифеста надстройки в эту `wef` папку.

    ![Папка Wef в Office для Mac.](../images/all-my-files.png)

1. Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.

1. В Word выберите **команду "Вставка** >  >  надстроек " Мои надстройки **"** (раскрывающееся меню), а затем выберите свою надстройку.

    ![Мои надстройки в Office на Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.

1. Проверьте, отображается ли ваша надстройка в Word.

    ![Надстройка Office, отображаемая в Office для Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Удаление неопубликоваемой надстройки

Вы можете удалить ранее загруженную неопубликованную надстройку, очистите кэш Office на компьютере. Сведения о том, как очистить кэш для каждой платформы и приложения, см. в статье "Очистка [кэша Office"](clear-cache.md).

## <a name="see-also"></a>См. также

- [Загрузка неопубликованных надстроек Office на iPad для тестирования](sideload-an-office-add-in-on-ipad.md)
- [Отладка надстроек Office на Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md)
