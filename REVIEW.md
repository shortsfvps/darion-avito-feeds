# REVIEW: Корректировка Title и Description под лимиты Авито

## Задача
Скорректировать логику формирования полей Title и Description в скрипте генерации XML-фидов для Авито под требования клиента и лимиты площадки (Title <= 50 символов с артикулом в скобках, Description со спецификациями и полным описанием из AH <= 5000 символов).

## Что сделано
1. В `DescriptionBuilder.build_description`:
   - Сохранена и структурирована очередность: сначала шапка (`name`, `sku`, `brand`) и блок `[ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ]`.
   - Под характеристиками выводится полное продающее описание из столбца AH (`extra_desc`) в блоке `[ОПИСАНИЕ И ПРЕИМУЩЕСТВА]` без обрезки абзацев.
   - В конце добавлен блок преимуществ компании «Дарион Свет» и контакты.
   - Реализован жесткий контроль лимита длины тега: если суммарный объем превышает 5 000 символов, строка безопасно усекается до 5 000 символов без висячих знаков.
2. В `AvitoFeedGenerator.format_title`:
   - Добавлен параметр `sku` для формирования заголовка вида `«[Название] ([Артикул])»`.
   - Заголовок гарантированно укладывается в `<= 50` символов.
   - Реализована очистка от старых концевых скобок с артикулами и обрезка базового наименования по точной границе слов без висячих знаков пунктуации.
3. В `AvitoFeedGenerator.build_ad_element`:
   - Вызов `format_title` обновлен: передаются `prod["name"]` и `prod["sku"]`.

## Список изменённых файлов
- `scripts/generate_feeds.py`
- `output/led_luminaires_feed.xml`
- `output/track_systems_feed.xml`
- `PLAN.md`
- `.gitignore`
- `.review-base`

## Изменённые фрагменты кода (diff)
```diff
--- a/scripts/generate_feeds.py
+++ b/scripts/generate_feeds.py
@@ -374,27 +374,13 @@ class DescriptionBuilder:
 
         specs_block = "\n".join(specs) if specs else "• Характеристики соответствуют паспорту изделия"
 
-        # Дополнительное описание
-        extra_desc = prod.get("extra_desc", "")
-        marketing_text = ""
-        if extra_desc and len(extra_desc) > 30:
-            paragraphs = [p.strip() for p in extra_desc.split("\n\n") if p.strip()]
-            meaningful = [
-                p for p in paragraphs
-                if p.lower() != name.lower() and not p.lower().startswith(name.lower()[:20])
-            ]
-            if meaningful:
-                summary_p = meaningful[0]
-                marketing_text = f"\n[ОПИСАНИЕ И ПРЕИМУЩЕСТВА]\n{summary_p}\n"
-
-        description_text = f"""{name}
-Артикул: {sku}
-Производитель: {brand}
-
-[ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ]
-{specs_block}
-{marketing_text}
-[ПРЕИМУЩЕСТВА И УСЛОВИЯ КОМПАНИИ «ДАРИОН СВЕТ»]
+        # Полное продающее описание из столбца AH (34) без обрезки текста
+        extra_desc = str(prod.get("extra_desc") or "").strip()
+        marketing_block = ""
+        if extra_desc:
+            marketing_block = f"\n\n[ОПИСАНИЕ И ПРЕИМУЩЕСТВА]\n{extra_desc}"
+
+        company_block = """[ПРЕИМУЩЕСТВА И УСЛОВИЯ КОМПАНИИ «ДАРИОН СВЕТ»]
 - Официальная гарантия производителя на всю светотехнику.
 - Быстрая отгрузка со склада в Санкт-Петербурге.
 - Работаем с юр. и физ. лицами (оплата по счету с НДС 20% и без НДС).
@@ -403,7 +389,15 @@ class DescriptionBuilder:
 
 Звоните или пишите в сообщения на Авито — рассчитаем освещенность объекта и подберем необходимое оборудование!"""
 
-        return description_text.strip()
+        header_block = f"{name}\nАртикул: {sku}\nПроизводитель: {brand}\n\n[ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ]\n{specs_block}"
+
+        description_text = f"{header_block}{marketing_block}\n\n{company_block}".strip()
+
+        # Системный контроль: суммарная длина тега Description <= 5000 символов (лимит Авито)
+        if len(description_text) > 5000:
+            description_text = description_text[:5000].rstrip(' ,.-;/')
+
+        return description_text
 
 
 class FeedRouter:
@@ -429,22 +423,33 @@ class AvitoFeedGenerator:
         self.feed_buckets = {k: [] for k in FEEDS_CONFIG.keys()}
 
     @staticmethod
-    def format_title(name: str, max_len: int = 50) -> str:
-        """Ограничивает длину заголовка для Авито (максимум 50 символов) без обрыва слова на полуслове."""
+    def format_title(name: str, sku: str = "", max_len: int = 50) -> str:
+        """
+        Формирует заголовок объявления для Авито в формате «[Название] ([Артикул])».
+        Общая длина строки СТРОГО <= max_len (по умолчанию 50 символов).
+        Базовое наименование аккуратно обрезается по границе слов без висячих знаков пунктуации.
+        """
         name = (name or "").strip()
-        if len(name) <= max_len:
-            return name
-        # Удаляем хвостовые скобки с артикулами (VRN-UNE-48-G40K67-U)
-        cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()
-        if len(cleaned) <= max_len:
-            return cleaned
-        # Обрезка по границе слова
-        last_space = cleaned.rfind(' ', 0, max_len)
-        if last_space > 0:
-            truncated = cleaned[:last_space].rstrip(' ,.-')
+        sku = (sku or "").strip()
+
+        # Удаляем хвостовые скобки с артикулами (VRN-UNE-48-G40K67-U) или упаковкой
+        cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip() or name
+        suffix = f" ({sku})" if sku else ""
+        avail = max_len - len(suffix)
+        if avail <= 0:
+            return (cleaned + suffix)[:max_len]
+
+        if len(cleaned) <= avail:
+            base = cleaned
         else:
-            truncated = cleaned[:max_len]
-        return truncated if len(truncated) <= max_len else truncated[:max_len]
+            if len(cleaned) > avail and cleaned[avail] == ' ':
+                base = cleaned[:avail].rstrip(' ,.-;/')
+            else:
+                sp = cleaned.rfind(' ', 0, avail)
+                base = cleaned[:sp].rstrip(' ,.-;/') if sp > 0 else cleaned[:avail].rstrip(' ,.-;/')
+
+        t = f"{base}{suffix}".strip()
+        return t[:max_len].rstrip(' ,.-;/') if len(t) > max_len else t
 
     def build_ad_element(self, prod: dict, feed_key: str) -> etree.Element:
         """Формирует XML-элемент <Ad> со всеми обязательными и категорийными тегами."""
@@ -458,7 +463,7 @@ class AvitoFeedGenerator:
         address.text = getattr(config, "DEFAULT_ADDRESS", "Санкт-Петербург, Студенческая ул., 10")
 
         title = etree.SubElement(ad, "Title")
-        title.text = self.format_title(prod["name"], max_len=50)
+        title.text = self.format_title(prod["name"], prod["sku"], max_len=50)
```

## Команда запуска и вывод валидации
Команда: `python scripts/generate_feeds.py`
Вывод:
```text
Всего проверено объявлений: 35
Максимальная длина Title среди всех: 50 (лимит <= 50)
Максимальная длина Description среди всех: 4901 (лимит <= 5000)
Товар ID: 14690612038169 -> Title (50 симв.): Панель светодиодная универсальная (14690612038169)
Товар ID: V1007          -> Title (43 симв.): Светодиодный светильник Virona 48Вт (V1007)
```

## Что не проверено или под вопросом
Все требования клиента и лимиты Авито проверены и подтверждены скриптами валидации. Открытых вопросов нет.
