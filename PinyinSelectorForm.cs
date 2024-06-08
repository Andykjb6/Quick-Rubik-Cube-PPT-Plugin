using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PpSelectionType = Microsoft.Office.Interop.PowerPoint.PpSelectionType;

namespace 课件帮PPT助手
{
    public partial class PinyinSelectorForm : Form
    {
        public PinyinSelectorForm()
        {
            InitializeComponent();
            this.FormClosing += (sender, e) => { ((Ribbon1)Globals.Ribbons.Ribbon1).PinyinSelectorFormClosed(); };
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            UpdateComboBoxOptions();
        }

        private void ReplaceButton_Click(object sender, EventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionText && comboBox.SelectedItem != null)
            {
                selection.TextRange2.Text = comboBox.SelectedItem.ToString();
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void UpdateComboBoxOptions()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
            {
                string selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.Text.Trim().ToLower();
                string[] pinyinOptions;

                if (selectedText == "yu")
                {
                    pinyinOptions = new string[] { "yū", "yú", "yǔ", "yù" };
                }
                else
                {
                    pinyinOptions = FindClosestPinyin(selectedText);
                }

                comboBox.DataSource = pinyinOptions;
            }
        }

        private string[] FindClosestPinyin(string inputText)
        {
            string[] pinyinLibrary = { "ā", "á", "ǎ", "à", "āi", "ái", "ǎi", "ài", "ān", "án", "ǎn", "àn", "ānɡ", "ánɡ", "ǎnɡ", "ànɡ", "āo", "áo", "ǎo", "ào", "bā", "bá", "bǎ", "bà", "bāi", "bái", "bǎi", "bài", "bān", "bán", "bǎn", "bàn", "bānɡ", "bánɡ", "bǎnɡ", "bànɡ", "bāo", "báo", "bǎo", "bào", "bēi", "béi", "běi", "bèi", "bēn", "bén", "běn", "bèn", "bēnɡ", "bénɡ", "běnɡ", "bènɡ", "bī", "bí", "bǐ", "bì", "biān", "bián", "biǎn", "biàn", "biāo", "biáo", "biǎo", "biào", "biē", "bié", "biě", "biè", "bīn", "bín", "bǐn", "bìn", "bīnɡ", "bínɡ", "bǐnɡ", "bìnɡ", "bō", "bó", "bǒ", "bò", "bū", "bú", "bǔ", "bù", "cā", "cá", "cǎ", "cà", "cāi", "cái", "cǎi", "cài", "cān", "cán", "cǎn", "càn", "cāng", "cáng", "cǎng", "càng", "cāo", "cáo", "cǎo", "cào", "cē", "cé", "cě", "cè", "cēn", "cén", "cěn", "cèn", "cēng", "céng", "cěng", "cèng", "chā", "chá", "chǎ", "chà", "chāi", "chái", "chǎi", "chài", "chān", "chán", "chǎn", "chàn", "chāng", "cháng", "chǎng", "chàng", "chāo", "cháo", "chǎo", "chào", "chē", "ché", "chě", "chè", "chēn", "chén", "chěn", "chèn", "chēng", "chéng", "chěng", "chèng", "chī", "chí", "chǐ", "chì", "chōng", "chóng", "chǒng", "chòng", "chōu", "chóu", "chǒu", "chòu", "chū", "chú", "chǔ", "chù", "chuā", "chuá", "chuǎ", "chuà", "chuāi", "chuái", "chuǎi", "chuài", "chuān", "chuán", "chuǎn", "chuàn", "chuāng", "chuáng", "chuǎng", "chuàng", "chuī", "chuí", "chuǐ", "chuì", "chūn", "chún", "chǔn", "chùn", "chuō", "chuó", "chuǒ", "chuò", "cī", "cí", "cǐ", "cì", "cōng", "cóng", "cǒng", "còng", "cōu", "cóu", "cǒu", "còu", "cū", "cú", "cǔ", "cù", "cuān", "cuán", "cuǎn", "cuàn", "cuī", "cuí", "cuǐ", "cuì", "cūn", "cún", "cǔn", "cùn", "cuō", "cuó", "cuǒ", "cuò", "dā", "dá", "dǎ", "dà", "dāi", "dái", "dǎi", "dài", "dān", "dán", "dǎn", "dàn", "dāng", "dáng", "dǎng", "dàng", "dāo", "dáo", "dǎo", "dào", "dē", "dé", "dě", "dè", "dēi", "déi", "děi", "dèi", "dēn", "dén", "děn", "dèn", "dēng", "déng", "děng", "dèng", "dī", "dí", "dǐ", "dì", "diān", "dián", "diǎn", "diàn", "diāo", "diáo", "diǎo", "diào", "diē", "dié", "diě", "diè", "dīng", "díng", "dǐng", "dìng", "diū", "diú", "diǔ", "diù", "dōng", "dóng", "dǒng", "dòng", "dōu", "dóu", "dǒu", "dòu", "dū", "dú", "dǔ", "dù", "duān", "duán", "duǎn", "duàn", "duī", "duí", "duǐ", "duì", "dūn", "dún", "dǔn", "dùn", "duō", "duó", "duǒ", "duò", "ē", "é", "ě", "è", "ēi", "éi", "ěi", "èi", "ēn", "én", "ěn", "èn", "ēnɡ", "énɡ", "ěnɡ", "ènɡ", "ēr", "ér", "ěr", "èr", "fā", "fá", "fǎ", "fà", "fān", "fán", "fǎn", "fàn", "fāng", "fáng", "fǎng", "fàng", "fēi", "féi", "fěi", "fèi", "fēn", "fén", "fěn", "fèn", "fēng", "féng", "fěng", "fèng", "fō", "fó", "fǒ", "fò", "fōu", "fóu", "fǒu", "fòu", "fū", "fú", "fǔ", "fù", "gā", "gá", "gǎ", "gà", "gāi", "gái", "gǎi", "gài", "gān", "gán", "gǎn", "gàn", "gāng", "gáng", "gǎng", "gàng", "gāo", "gáo", "gǎo", "gào", "gē", "gé", "gě", "gè", "gēi", "géi", "gěi", "gèi", "gēn", "gén", "gěn", "gèn", "gēng", "géng", "gěng", "gèng", "gōng", "góng", "gǒng", "gòng", "gōu", "góu", "gǒu", "gòu", "gū", "gú", "gǔ", "gù", "guā", "guá", "guǎ", "guà", "guāi", "guái", "guǎi", "guài", "guān", "guán", "guǎn", "guàn", "guāng", "guáng", "guǎng", "guàng", "guī", "guí", "guǐ", "guì", "gūn", "gún", "gǔn", "gùn", "guō", "guó", "guǒ", "gùo", "hā", "há", "hǎ", "hà", "hāi", "hái", "hǎi", "hài", "hān", "hán", "hǎn", "hàn", "hāng", "háng", "hǎng", "hàng", "hāo", "háo", "hǎo", "hào", "hē", "hé", "hě", "hè", "hēi", "héi", "hěi", "hèi", "hēn", "hén", "hěn", "hèn", "hēng", "héng", "hěng", "hèng", "hōng", "hóng", "hǒng", "hòng", "hōu", "hóu", "hǒu", "hòu", "hū", "hú", "hǔ", "hù", "huā", "huá", "huǎ", "huà", "huāi", "huái", "huǎi", "huài", "huān", "huán", "huǎn", "huàn", "huāng", "huáng", "huǎng", "huàng", "huī", "huí", "huǐ", "huì", "hūn", "hún", "hǔn", "hùn", "huō", "huó", "huǒ", "huò", "ī", "í", "ǐ", "ì", "iē", "ié", "iě", "iè", "īn", "ín", "ǐn", "ìn", "īnɡ", "ínɡ", "ǐnɡ", "ìnɡ", "iū", "iú", "iǔ", "iù", "jī", "jí", "jǐ", "jì", "jiā", "jiá", "jiǎ", "jià", "jiāi", "jiái", "jiǎi", "jiài", "jiān", "jián", "jiǎn", "jiàn", "jiāng", "jiáng", "jiǎng", "jiàng", "jiāo", "jiáo", "jiǎo", "jiào", "jiē", "jié", "jiě", "jiè", "jīn", "jín", "jǐn", "jìn", "jīng", "jíng", "jǐng", "jìng", "jiōng", "jióng", "jiǒng", "jiòng", "jiū", "jiú", "jiǔ", "jiù", "jū", "jú", "jǔ", "jù", "juān", "juán", "juǎn", "juàn", "juē", "jué", "juě", "juè", "jūn", "jún", "jǔn", "jùn", "kā", "ká", "kǎ", "kà", "kāi", "kái", "kǎi", "kài", "kān", "kán", "kǎn", "kàn", "kāng", "káng", "kǎng", "kàng", "kāo", "káo", "kǎo", "kào", "kē", "ké", "kě", "kè", "kēn", "kén", "kěn", "kèn", "kēng", "kéng", "kěng", "kèng", "kōng", "kóng", "kǒng", "kòng", "kōu", "kóu", "kǒu", "kòu", "kū", "kú", "kǔ", "kù", "kuā", "kuá", "kuǎ", "kuà", "kuāi", "kuái", "kuǎi", "kuài", "kuān", "kuán", "kuǎn", "kuàn", "kuāng", "kuáng", "kuǎng", "kuàng", "kuī", "kuí", "kuǐ", "kuì", "kūn", "kún", "kǔn", "kùn", "kuō", "kuó", "kuǒ", "kuò", "lā", "lá", "lǎ", "là", "laī", "laí", "lǎi", "laì", "lān", "lán", "lǎn", "làn", "lāng", "láng", "lǎng", "làng", "lāo", "láo", "lǎo", "lào", "lē", "lé", "lě", "lè", "lēi", "léi", "lěi", "lèi", "lēng", "léng", "lěng", "lèng", "lī", "lí", "lǐ", "lì", "liān", "lián", "liǎn", "liàn", "liāng", "liáng", "liǎng", "liàng", "liāo", "liáo", "liǎo", "liào", "liē", "lié", "liě", "liè", "līn", "lín", "lǐn", "lìn", "līng", "líng", "lǐng", "lìng", "liū", "liú", "liǔ", "liù", "lōng", "lóng", "lǒng", "lòng", "lōu", "lóu", "lǒu", "lòu", "lū", "lú", "lǔ", "lù", "luān", "luán", "luǎn", "luàn", "luē", "lué", "luě", "luè", "lūn", "lún", "lǔn", "lùn", "luō", "luó", "luǒ", "luò", "mā", "má", "mǎ", "mà", "maī", "maí", "mǎi", "maì", "mān", "mán", "mǎn", "màn", "māng", "máng", "mǎng", "màng", "māo", "máo", "mǎo", "mào", "mē", "mé", "mě", "mè", "mēi", "méi", "měi", "mèi", "mēn", "mén", "měn", "mèn", "mēng", "méng", "měng", "mèng", "mī", "mí", "mǐ", "mì", "miān", "mián", "miǎn", "miàn", "miāo", "miáo", "miǎo", "miào", "miē", "mié", "miě", "miè", "mīn", "mín", "mǐn", "mìn", "mīng", "míng", "mǐng", "mìng", "miū", "miú", "miǔ", "miù", "mō", "mó", "mǒ", "mò", "mōu", "móu", "mǒu", "mòu", "mū", "mú", "mǔ", "mù", "nā", "ná", "nǎ", "nà", "naī", "naí", "nǎi", "naì", "nān", "nán", "nǎn", "nàn", "nāng", "náng", "nǎng", "nàng", "nāo", "náo", "nǎo", "nào", "nē", "né", "ně", "nè", "neī", "neí", "něi", "neì", "nēn", "nén", "něn", "nèn", "nēng", "néng", "něng", "nèng", "nī", "ní", "nǐ", "nì", "niān", "nián", "niǎn", "niàn", "niāng", "niáng", "niǎng", "niàng", "niāo", "niáo", "niǎo", "niào", "niē", "nié", "niě", "niè", "nīn", "nín", "nǐn", "nìn", "nīng", "níng", "nǐng", "nìng", "niū", "niú", "niǔ", "niù", "nōng", "nóng", "nǒng", "nòng", "nū", "nú", "nǔ", "nù", "nuān", "nuán", "nuǎn", "nuàn", "nuē", "nué", "nuě", "nuè", "nuō", "nuó", "nuǒ", "nuò", "nǘ", "nǚ", "nǜ", "nǘè", "ō", "ó", "ǒ", "ò", "ōnɡ", "ónɡ", "ǒnɡ", "ònɡ", "ōu", "óu", "ǒu", "òu", "pā", "pá", "pǎ", "pà", "pāi", "pái", "pǎi", "pài", "pān", "pán", "pǎn", "pàn", "pāng", "páng", "pǎng", "pàng", "pāo", "páo", "pǎo", "pào", "pēi", "péi", "pěi", "pèi", "pēn", "pén", "pěn", "pèn", "pēng", "péng", "pěng", "pèng", "pī", "pí", "pǐ", "pì", "piān", "pián", "piǎn", "piàn", "piāo", "piáo", "piǎo", "piào", "piē", "pié", "piě", "piè", "pīn", "pín", "pǐn", "pìn", "pīng", "píng", "pǐng", "pìng", "pō", "pó", "pǒ", "pò", "pōu", "póu", "pǒu", "pòu", "pū", "pú", "pǔ", "pù", "qī", "qí", "qǐ", "qì", "qiā", "qiá", "qǐa", "qìa", "qiān", "qián", "qǐan", "qìan", "qiāng", "qiáng", "qǐang", "qìang", "qiāo", "qiáo", "qǐao", "qìao", "qiē", "qié", "qǐe", "qìe", "qīn", "qín", "qǐn", "qìn", "qīng", "qíng", "qǐng", "qìng", "qiōnɡ", "qiónɡ", "qiǒnɡ", "qiònɡ", "qiū", "qiú", "qiǔ", "qiù", "qū", "qú", "qǔ", "qù", "quān", "quán", "quǎn", "quàn", "quē", "qué", "quě", "què", "qūn", "qún", "qǔn", "qùn", "rān", "rán", "rǎn", "ràn", "rāng", "ráng", "rǎng", "ràng", "rāo", "ráo", "rǎo", "rào", "rē", "ré", "rě", "rè", "rēn", "rén", "rěn", "rèn", "rēng", "réng", "rěng", "rèng", "rī", "rí", "rǐ", "rì", "rōng", "róng", "rǒng", "ròng", "rōu", "róu", "rǒu", "ròu", "rū", "rú", "rǔ", "rù", "ruān", "ruán", "ruǎn", "ruàn", "ruī", "ruí", "ruǐ", "ruì", "rūn", "rún", "rǔn", "rùn", "ruō", "ruó", "ruǒ", "ruò", "sā", "sá", "sǎ", "sà", "sāi", "sái", "sǎi", "sài", "sān", "sán", "sǎn", "sàn", "sāng", "sáng", "sǎng", "sàng", "sāo", "sáo", "sǎo", "sào", "sē", "sé", "sě", "sè", "sēn", "sén", "sěn", "sèn", "sēng", "séng", "sěng", "sèng", "shā", "shá", "shǎ", "shà", "shāi", "shái", "shǎi", "shài", "shān", "shán", "shǎn", "shàn", "shāng", "sháng", "shǎng", "shàng", "shāo", "sháo", "shǎo", "shào", "shē", "shé", "shě", "shè", "shēi", "shéi", "shěi", "shèi", "shēn", "shén", "shěn", "shèn", "shēng", "shéng", "shěng", "shèng", "shī", "shí", "shǐ", "shì", "shōu", "shóu", "shǒu", "shòu", "shū", "shú", "shǔ", "shù", "shuā", "shuá", "shuǎ", "shuà", "shuāi", "shuái", "shuǎi", "shuài", "shuān", "shuán", "shuǎn", "shuàn", "shuāng", "shuáng", "shuǎng", "shuàng", "shuī", "shuí", "shuǐ", "shuì", "shu?n", "shú", "shǔn", "shùn", "shuō", "shuó", "shuǒ", "shuò", "sī", "sí", "sǐ", "sì", "sōng", "sóng", "sǒng", "sòng", "sū", "sú", "sǔ", "sù", "suān", "suán", "suǎn", "suàn", "suī", "suí", "suǐ", "suì", "su?n", "sún", "sǔn", "sùn", "suō", "suó", "suǒ", "suò", "tā", "tá", "tǎ", "tà", "tāi", "tái", "tǎi", "tài", "tān", "tán", "tǎn", "tàn", "tāng", "táng", "tǎng", "tàng", "tāo", "táo", "tǎo", "tào", "tē", "té", "tě", "tè", "tēng", "téng", "těng", "tèng", "tī", "tí", "tǐ", "tì", "tiān", "tián", "tiǎn", "tiàn", "tiāo", "tiáo", "tiǎo", "tiào", "tiē", "tié", "tiě", "tiè", "tīng", "tíng", "tǐng", "tìng", "tōng", "tóng", "tǒng", "tòng", "tōu", "tóu", "tǒu", "tòu", "tū", "tú", "tǔ", "tù", "tuān", "tuán", "tuǎn", "tuàn", "tuī", "tuí", "tuǐ", "tuì", "tūn", "tún", "tǔn", "tùn", "tuō", "tuó", "tuǒ", "tuò", "ū", "ú", "ǔ", "ù", "uē", "ué", "uě", "uè", "uī", "uí", "uǐ", "uì", "ūn", "ún", "ǔn", "ùn", "ǖ", "ǘ", "ǚ", "ǜ", "ǖn", "ǘn", "ǚn", "ǜn", "wā", "wá", "wǎ", "wà", "wāi", "wái", "wǎi", "wài", "wān", "wán", "wǎn", "wàn", "wāng", "wáng", "wǎng", "wàng", "wēi", "wéi", "wěi", "wèi", "wēn", "wén", "wěn", "wèn", "wēng", "wéng", "wěng", "wèng", "wō", "wó", "wǒ", "wò", "wū", "wú", "wǔ", "wù", "xī", "xí", "xǐ", "xì", "xiā", "xiá", "xiǎ", "xià", "xiān", "xián", "xiǎn", "xiàn", "xiāng", "xiáng", "xiǎng", "xiàng", "xiāo", "xiáo", "xiǎo", "xiào", "xiē", "xié", "xiě", "xiè", "xīn", "xín", "xǐn", "xìn", "xīng", "xíng", "xǐng", "xìng", "xiōng", "xióng", "xiǒng", "xiòng", "xiū", "xiú", "xiǔ", "xiù", "xū", "xú", "xǔ", "xù", "xuān", "xuán", "xuǎn", "xuàn", "xuē", "xué", "xuě", "xuè", "xūn", "xún", "xǔn", "xùn", "yā", "yá", "yǎ", "yà", "yān", "yán", "yǎn", "yàn", "yāng", "yáng", "yǎng", "yàng", "yāo", "yáo", "yǎo", "yào", "yē", "yé", "yě", "yè", "yī", "yí", "yǐ", "yì", "yīn", "yín", "yǐn", "yìn", "yīng", "yíng", "yǐng", "yìng", "yō", "yó", "yǒ", "yò", "yōng", "yóng", "yǒng", "yòng", "yōu", "yóu", "yǒu", "yòu", "yū", "yú", "yǔ", "yù", "yuān", "yuán", "yuǎn", "yuàn", "yuē", "yué", "yuě", "yuè", "yūn", "yún", "yǔn", "yùn", "zā", "zá", "zǎ", "zà", "zāi", "zái", "zǎi", "zài", "zān", "zán", "zǎn", "zàn", "zāng", "záng", "zǎng", "zàng", "zāo", "záo", "zǎo", "zào", "zē", "zé", "zě", "zè", "zēi", "zéi", "zěi", "zèi", "zēn", "zén", "zěn", "zèn", "zēng", "zéng", "zěng", "zèng", "zhā", "zhá", "zhǎ", "zhà", "zhāi", "zhái", "zhǎi", "zhài", "zhān", "zhán", "zhǎn", "zhàn", "zhāng", "zháng", "zhǎng", "zhàng", "zhāo", "zháo", "zhǎo", "zhào", "zhē", "zhé", "zhě", "zhè", "zhēi", "zhéi", "zhěi", "zhèi", "zhēn", "zhén", "zhěn", "zhèn", "zhēng", "zhéng", "zhěng", "zhèng", "zhī", "zhí", "zhǐ", "zhì", "zhōng", "zhóng", "zhǒng", "zhòng", "zhōu", "zhóu", "zhǒu", "zhòu", "zhū", "zhú", "zhǔ", "zhù", "zhuā", "zhuá", "zhuǎ", "zhuà", "zhuāi", "zhuái", "zhuǎi", "zhuài", "zhuān", "zhuán", "zhuǎn", "zhuàn", "zhuānɡ", "zhuánɡ", "zhuǎnɡ", "zhuàng", "zhuī", "zhuí", "zhuǐ", "zhuì", "zhūn", "zhún", "zhǔn", "zhùn", "zī", "zí", "zǐ", "zì", "zōng", "zóng", "zǒng", "zòng", "zōu", "zóu", "zǒu", "zòu", "zū", "zú", "zǔ", "zù", "zuān", "zuán", "zuǎn", "zuàn", "zuī", "zuí", "zuǐ", "zuì", "zūn", "zún", "zǔn", "zùn", "zuō", "zuó", "zuǒ", "zuò", };
            return pinyinLibrary.Where(p => NormalizePinyin(p).StartsWith(NormalizePinyin(inputText))).Take(4).ToArray();
        }

        private string NormalizePinyin(string pinyin)
        {
            System.Collections.Generic.Dictionary<char, char> accentMapping = new System.Collections.Generic.Dictionary<char, char>
            {
                { 'ā', 'a' }, { 'á', 'a' }, { 'ǎ', 'a' }, { 'à', 'a' },
                { 'ō', 'o' }, { 'ó', 'o' }, { 'ǒ', 'o' }, { 'ò', 'o' },
                { 'ē', 'e' }, { 'é', 'e' }, { 'ě', 'e' }, { 'è', 'e' },
                { 'ī', 'i' }, { 'í', 'i' }, { 'ǐ', 'i' }, { 'ì', 'i' },
                { 'ū', 'u' }, { 'ú', 'u' }, { 'ǔ', 'u' }, { 'ù', 'u' },
                { 'ǖ', 'ü' }, { 'ǘ', 'ü' }, { 'ǚ', 'ü' }, { 'ǜ', 'ü' }
            };
            return new string(pinyin.Select(c => accentMapping.ContainsKey(c) ? accentMapping[c] : c).ToArray());
        }
    }
}
