# <a name="dialog-boxes-in-office-add-ins"></a>Диалоговые окна в надстройках Office
 
Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.

**Пример диалогового окна**

![Изображение, на котором показан типичный макет диалогового окна](../images/overview_withApp_dialog.png)

### <a name="best-practices"></a>Рекомендации

|**Рекомендуется**|**Не рекомендуется**|
|:-----|:--------|
|<ul><li>Укажите описательное название, содержащее имя надстройки и название текущей задачи.</li></ul>|<ul><li>Не включайте в него название вашей компании.</li></ul>|
||<ul><li>Не открывайте диалоговое окно, если этого не требует сценарий.</li></ul>|

## <a name="implementation"></a>Реализация

Пример реализации диалогового окна см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Пример конструктивного шаблона](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [Ресурсы для разработки на сайте GitHub](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Объект Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog)


