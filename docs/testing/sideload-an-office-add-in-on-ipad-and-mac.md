---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: Проверьте Office надстройку на iPad Mac с помощью боковой загрузки.
ms.date: 09/02/2020
ms.localizationpriority: medium
ms.openlocfilehash: 04609f8cceee20403c25ec91a8ca75adf82b51c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154949"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования

Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.

## <a name="prerequisites-for-office-on-ios"></a>Предварительные требования (Office для iOS)

- Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).
  > [!IMPORTANT]
  > Если вы используете macOS Catalina, [iTunes](https://support.apple.com/HT210200) больше не доступен, поэтому следует следовать инструкциям в разделе Sideload надстройки на Excel или Word на iPad с помощью [macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) позже в этой статье.

- Установлен iPad iOS 8.2 или более [](https://apps.apple.com/app/microsoft-excel/id586683407) поздней Excel [или Word,](https://apps.apple.com/app/microsoft-word/id586447913) а также синхронизированный кабель.

- XML-файл манифеста для надстройки, которую вы хотите протестировать.

## <a name="prerequisites-for-office-on-mac"></a>Предварительные требования (Office для Mac)

- Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).

- Word для Mac версии 15.18 (160109).

- Excel для Mac версии 15.19 (160206).

- PowerPoint для Mac версии 15.24 (160614)

- XML-файл манифеста для надстройки, которую вы хотите протестировать.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Sideload an add-in on Excel word on iPad using iTunes

1. Подключите iPad к компьютеру с помощью кабеля для синхронизации. Если вы впервые подключите iPad к компьютеру, вам будет предложено использовать **Trust This Computer?**. Выберите **Доверять**.

2. В iTunes под строкой меню выберите значок **iPad**.

3. В левой части iTunes в разделе **Параметры** выберите **Приложения**.

4. В правой части iTunes прокрутите окно вниз до раздела **Общий доступ к файлам**, а затем в столбце **Надстройки** выберите **Excel** или **Word**.

5. В нижней части **столбца Excel** **или Word Documents** выберите Добавить **файл,** а затем выберите файл манифеста .xml надстройки, необходимой для загрузки.

6. Откройте приложение Excel или Word на iPad. Если приложение Excel Word уже запущено, выберите кнопку **Главная,** а затем закрой и перезапустите приложение.

7. Откройте документ.

8. Выберите **надстройки** на вкладке **Вставка.** (На вкладке **Вставить** может потребоваться прокрутка по горизонтали, пока не увидите кнопку **Надстройки.)** Ваша надстройка с боковой загрузкой доступна для вставки под заголовком **Developer** в пользовательском интерфейсе **надстройки.**

    ![Вставьте надстройки в Excel приложении.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Sideload an add-in on Excel word on iPad using macOS Catalina

> [!IMPORTANT]
> С введением macOS Catalina Apple прекратила [iTunes](https://support.apple.com/HT210200) на Mac и интегрированные функции, необходимые для загрузки приложений в **Finder**.

1. Подключите iPad к компьютеру с помощью кабеля для синхронизации. Если вы впервые подключите iPad к компьютеру, вам будет предложено использовать **Trust This Computer?**. Выберите **Доверять**. Вы также можете быть заданы вопросы, если это новый iPad или если вы восстанавливаете один.

2. В Finder в **статье Locations** выберите значок **iPad** ниже панели меню.

3. В верхней части окна Finder нажмите кнопку **Файлы,** а затем **найдите** Excel **или Word**.

4. Из другого окна Finder перетащите и manifest.xml файл надстройки, который необходимо загрузить в файл **Excel** **Word** в первом окне Finder.

5. Откройте приложение Excel или Word на iPad. Если приложение Excel Word уже запущено, выберите кнопку **Главная,** а затем закрой и перезапустите приложение.

6. Откройте документ.

7. Выберите **надстройки** на вкладке **Вставка.** (На вкладке **Вставить** может потребоваться прокрутка по горизонтали, пока не увидите кнопку **Надстройки.)** Ваша надстройка с боковой загрузкой доступна для вставки под заголовком **Developer** в пользовательском интерфейсе **надстройки.**

    ![Вставьте надстройки в Excel приложении.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Загрузка неопубликованной надстройки в Office для Mac

> [!NOTE]
> Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).

1. Откройте **терминал** и перейдите в одну из следующих папок, где вы сохраните файл манифеста надстройки. Если папки `wef` нет на компьютере, создайте ее.

    - Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Откройте папку в **Finder** с помощью команды `open .` (включая период или точку). Скопируйте файл манифеста надстройки в эту папку.

    ![Папка Wef в Office на Mac.](../images/all-my-files.png)

3. Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.

4. В Word выберите **вставьте** надстройки Мои надстройки (выпадаемое меню), а затем  >    >   выберите надстройки.

    ![Мои надстройки в Office на Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.

5. Проверьте, отображается ли ваша надстройка в Word.

    ![Office Надстройка, отображаемая в Office mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Удаление боковой надстройки

Вы можете удалить ранее загруженную надстройку, очищая кэш Office на компьютере. Сведения о том, как очистить кэш для каждой платформы и приложения, можно найти в статье [Clear the Office кэш.](clear-cache.md)

## <a name="see-also"></a>См. также

- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)
