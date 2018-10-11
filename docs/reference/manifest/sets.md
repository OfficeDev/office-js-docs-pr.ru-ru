# <a name="sets-element"></a>Элемент Sets

Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.

**Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Содержится в

[Требования](requirements.md)

## <a name="can-contain"></a>Может содержать

[Множество](set.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](set.md). Значение по умолчанию: "1.1".|

## <a name="remarks"></a>Замечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии и наборы обязательных элементов Office](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

