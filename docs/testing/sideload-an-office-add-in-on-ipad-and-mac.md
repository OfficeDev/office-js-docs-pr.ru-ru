---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: ''
ms.date: 11/26/2019
localization_priority: Priority
ms.openlocfilehash: 4187755ecb806d5034efb67e022f0fe91bbd2559
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629744"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования

Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать. 

## <a name="prerequisites-for-office-on-ios"></a>Предварительные требования (Office для iOS)

- Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).
    
- iPad под управлением iOS 8.2 или более поздней версии, на котором установлено приложение [Excel на iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) и к которому подключен кабель для синхронизации.
    
- XML-файл манифеста для надстройки, которую вы хотите протестировать.
    

## <a name="prerequisites-for-office-on-mac"></a>Предварительные требования (Office для Mac)

- Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).
    
- Word для Mac версии 15.18 (160109).
   
- Excel для Mac версии 15.19 (160206).

- PowerPoint для Mac версии 15.24 (160614)
    
- XML-файл манифеста для надстройки, которую вы хотите протестировать.
    

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a>Загрузка неопубликованной надстройки в Excel или Word на iPad

1. Подключите iPad к компьютеру с помощью кабеля для синхронизации. Если вы подключаете iPad к компьютеру в первый раз, появится запрос **Доверять этому компьютеру?**. Выберите **Доверять**.

2. В iTunes под строкой меню выберите значок **iPad**.

3. В левой части iTunes в разделе  **Параметры** выберите **Приложения**.

4. В правой части iTunes прокрутите окно вниз до раздела  **Общий доступ к файлам**, а затем в столбце  **Надстройки** выберите **Excel** или **Word**.

5. В нижней части столбца  **Документы Excel** или **Документы Word** выберите элемент **Добавить файл**, а затем выберите XML-файл манифеста для надстройки, которую необходимо загрузить. 
    
6. Откройте приложение Excel или Word на iPad. Если приложение Excel или Word уже запущено, нажмите кнопку **Главная**, а затем закройте и перезапустите его.
    
7. Откройте документ.
    
8. На вкладке  **Вставка** выберите **Надстройки**. Загруженную надстройка можно добавить в разделе  **Разработчик** в пользовательском интерфейсе **Надстройки**.
    
    ![Вставка надстроек в приложение Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-in-office-on-mac"></a>Загрузка неопубликованной надстройки в Office для Mac

> [!NOTE]
> Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing).

1. Откройте **Терминал** и перейдите в одну из указанных ниже папок, чтобы сохранить в нее файл манифеста надстройки. Если папки `wef` нет на компьютере, создайте ее.
    
    - Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`    
    - Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
    
2. Откройте папку в **Finder** с помощью команды `open .` (включая точку). Скопируйте файл манифеста надстройки в эту папку.
    
    ![Папка Wef в Office для Mac](../images/all-my-files.png)

3. Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.
    
4. В Word выберите элементы **Вставка**  >  **Надстройки**  >  **Мои надстройки**, а затем выберите свою надстройку.
    
    ![Мои надстройки в Office для Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню. 
    
5. Проверьте, отображается ли ваша надстройка в Word.
    
    ![Надстройка в Office для Mac](../images/lorem-ipsum-wikipedia.png)
    
### <a name="clearing-the-office-applications-cache-on-a-mac"></a>Очистка кэша приложения Office на компьютере Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="see-also"></a>См. также

- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)
