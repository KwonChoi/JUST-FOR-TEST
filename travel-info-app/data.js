// Sample data for demo countries
const countriesData = [
    {
        id: 'japan',
        name: '日本',
        nameEn: 'Japan',
        region: 'asia',
        flag: 'https://flagcdn.com/w320/jp.jpg',
        capital: '东京',
        population: '1.26亿',
        area: '37.8万平方公里',
        highlight: '樱花之国，传统文化与现代科技完美融合',

        // 签证信息
        visa: {
            type: '旅游签证',
            duration: '15天',
            difficulty: '中等',
            requirements: [
                '护照原件（有效期6个月以上）',
                '签证申请表',
                '2寸白底照片2张',
                '在职证明及公司营业执照',
                '银行流水（近6个月）',
                '往返机票预订单',
                '酒店预订单'
            ],
            tips: '建议提前1-2个月申请，可通过旅行社代办简化材料'
        },

        // 货币信息
        currency: {
            name: '日元',
            code: 'JPY',
            symbol: '¥',
            rate: '1日元 ≈ 0.048人民币',
            paymentTips: '建议携带部分现金，大城市信用卡普及率高，小城市和传统市场多用现金'
        },

        // 语言信息
        language: {
            official: '日语',
            commonGreetings: [
                { chinese: '你好', japanese: 'こんにちは (Konnichiwa)' },
                { chinese: '谢谢', japanese: 'ありがとう (Arigatou)' },
                { chinese: '对不起', japanese: 'すみません (Sumimasen)' },
                { chinese: '再见', japanese: 'さようなら (Sayounara)' }
            ],
            usefulPhrases: [
                { chinese: '多少钱？', japanese: 'いくらですか？ (Ikura desu ka?)' },
                { chinese: '我不会说日语', japanese: '日本語が話せません (Nihongo ga hanasemasen)' },
                { chinese: '厕所在哪里？', japanese: 'トイレはどこですか？ (Toire wa doko desu ka?)' }
            ]
        },

        // 习俗禁忌
        customs: [
            {
                title: '礼仪文化',
                content: '见面要鞠躬，递接物品用双手，进入榻榻米房间要脱鞋'
            },
            {
                title: '用餐礼仪',
                content: '吃面时发出声音表示美味，不要将筷子插在饭中，传递食物不要用筷子直接传递'
            },
            {
                title: '公共行为',
                content: '地铁内保持安静，垃圾分类严格，温泉入浴前要先洗净身体'
            }
        ],

        // 玩法路线
        routes: [
            {
                title: '经典关西5日游',
                duration: '5天4夜',
                highlights: ['大阪城', '京都金阁寺', '奈良公园', '神户牛肉'],
                description: '体验日本传统文化，品尝地道美食，适合首次赴日游客'
            },
            {
                title: '东京深度7日游',
                duration: '7天6夜',
                highlights: ['浅草寺', '东京塔', '涩谷十字路口', '秋叶原', '镰仓'],
                description: '现代都市与传统文化的完美结合，购物美食一网打尽'
            }
        ],

        // 高频词汇
        vocabulary: [
            { chinese: '寿司', english: 'Sushi' },
            { chinese: '拉面', english: 'Ramen' },
            { chinese: '温泉', english: 'Onsen' },
            { chinese: '新干线', english: 'Shinkansen' },
            { chinese: '便利店', english: 'Convenience Store' },
            { chinese: '地铁', english: 'Subway/Metro' }
        ]
    },

    {
        id: 'kenya',
        name: '肯尼亚',
        nameEn: 'Kenya',
        region: 'africa',
        flag: 'https://flagcdn.com/w320/ke.jpg',
        capital: '内罗毕',
        population: '5300万',
        area: '58万平方公里',
        highlight: '非洲野生动物天堂，马赛马拉大迁徙的故乡',

        // 签证信息
        visa: {
            type: '电子签证/落地签',
            duration: '90天',
            difficulty: '简单',
            requirements: [
                '护照原件（有效期6个月以上）',
                '黄热病疫苗接种证明',
                '往返机票',
                '酒店预订单',
                '签证费50美元'
            ],
            tips: '建议提前申请电子签证，落地签排队时间较长'
        },

        // 货币信息
        currency: {
            name: '肯尼亚先令',
            code: 'KES',
            symbol: 'KSh',
            rate: '1肯尼亚先令 ≈ 0.055人民币',
            paymentTips: '建议携带美元现金兑换，大城市可使用信用卡，safari营地多为现金交易'
        },

        // 语言信息
        language: {
            official: '斯瓦希里语、英语',
            commonGreetings: [
                { chinese: '你好', swahili: 'Habari/Jambo' },
                { chinese: '谢谢', swahili: 'Asante' },
                { chinese: '再见', swahili: 'Kwaheri' }
            ],
            usefulPhrases: [
                { chinese: '多少钱？', english: 'How much?' },
                { chinese: '厕所在哪里？', english: 'Where is the toilet?' },
                { chinese: '救命！', english: 'Help!' }
            ]
        },

        // 习俗禁忌
        customs: [
            {
                title: '文化尊重',
                content: '拍摄当地人前要先征得同意，尊重马赛族等土著民族的传统'
            },
            {
                title: '安全提示',
                content: '野生动物园内禁止下车，晚上避免单独外出，注意饮食卫生'
            },
            {
                title: '环保意识',
                content: '禁止乱扔垃圾，不要购买野生动物制品，保护自然环境'
            }
        ],

        // 玩法路线
        routes: [
            {
                title: '经典野生动物safari7日游',
                duration: '7天6夜',
                highlights: ['马赛马拉国家保护区', '纳库鲁湖', '安博塞利国家公园', '内罗毕长颈鹿公园'],
                description: '观赏非洲五霸，体验壮观动物大迁徙，入住特色帐篷营地'
            },
            {
                title: '海滨度假5日游',
                duration: '5天4夜',
                highlights: ['蒙巴萨海滩', '迪亚尼海滩', '古城观光', '水上运动'],
                description: '印度洋海滨度假，享受阳光沙滩，体验斯瓦希里文化'
            }
        ],

        // 高频词汇（东非野生动物专业词汇）
        vocabulary: [
            { chinese: '狮子', english: 'Lion' },
            { chinese: '大象', english: 'Elephant' },
            { chinese: '犀牛', english: 'Rhino' },
            { chinese: '水牛', english: 'Buffalo' },
            { chinese: '豹子', english: 'Leopard' },
            { chinese: '长颈鹿', english: 'Giraffe' },
            { chinese: '斑马', english: 'Zebra' },
            { chinese: '角马', english: 'Wildebeest' },
            { chinese: '猎豹', english: 'Cheetah' },
            { chinese: '鬣狗', english: 'Hyena' },
            { chinese: 'safari', english: 'Safari' },
            { chinese: '国家公园', english: 'National Park' }
        ]
    },

    {
        id: 'france',
        name: '法国',
        nameEn: 'France',
        region: 'europe',
        flag: 'https://flagcdn.com/w320/fr.jpg',
        capital: '巴黎',
        population: '6700万',
        area: '55万平方公里',
        highlight: '浪漫之都，艺术与时尚的发源地',

        // 签证信息
        visa: {
            type: '申根签证',
            duration: '90天（180天内）',
            difficulty: '中等',
            requirements: [
                '护照原件（有效期6个月以上）',
                '签证申请表及照片',
                '旅行保险（3万欧元保额）',
                '在职证明及银行流水',
                '行程单及酒店预订单',
                '往返机票预订单'
            ],
            tips: '建议提前2-3个月申请，首次申请建议选择停留时间最长的国家'
        },

        // 货币信息
        currency: {
            name: '欧元',
            code: 'EUR',
            symbol: '€',
            rate: '1欧元 ≈ 7.8人民币',
            paymentTips: '信用卡普及率很高，建议携带国际信用卡，小费文化不盛行'
        },

        // 语言信息
        language: {
            official: '法语',
            commonGreetings: [
                { chinese: '你好', french: 'Bonjour' },
                { chinese: '谢谢', french: 'Merci' },
                { chinese: '对不起', french: 'Pardon' },
                { chinese: '再见', french: 'Au revoir' }
            ],
            usefulPhrases: [
                { chinese: '你会说英语吗？', french: 'Parlez-vous anglais?' },
                { chinese: '多少钱？', french: 'Combien ça coûte?' },
                { chinese: '厕所在哪里？', french: 'Où sont les toilettes?' }
            ]
        },

        // 习俗禁忌
        customs: [
            {
                title: '用餐礼仪',
                content: '用餐时双手放在桌上，面包直接放在桌布上，品酒时要先观察再品尝'
            },
            {
                title: '社交礼仪',
                content: '见面行贴面礼，称呼要用Monsieur/Madame，守时很重要'
            },
            {
                title: '文化尊重',
                content: '博物馆内禁止拍照，教堂内要保持安静，尊重当地文化传统'
            }
        ],

        // 玩法路线
        routes: [
            {
                title: '经典法国8日游',
                duration: '8天7夜',
                highlights: ['埃菲尔铁塔', '卢浮宫', '凡尔赛宫', '香榭丽舍大街', '蒙马特高地'],
                description: '巴黎经典景点全覆盖，感受浪漫之都的魅力'
            },
            {
                title: '南法薰衣草10日游',
                duration: '10天9夜',
                highlights: ['普罗旺斯', '阿维尼翁', '尼斯', '戛纳', '摩纳哥'],
                description: '南法风情之旅，薰衣草花田，蔚蓝海岸，感受法式乡村生活'
            }
        ],

        // 高频词汇
        vocabulary: [
            { chinese: '卢浮宫', english: 'Louvre' },
            { chinese: '埃菲尔铁塔', english: 'Eiffel Tower' },
            { chinese: '香榭丽舍大街', english: 'Champs-Élysées' },
            { chinese: '咖啡', english: 'Café' },
            { chinese: '红酒', english: 'Vin' },
            { chinese: '奶酪', english: 'Fromage' }
        ]
    },

    {
        id: 'thailand',
        name: '泰国',
        nameEn: 'Thailand',
        region: 'asia',
        flag: 'https://flagcdn.com/w320/th.jpg',
        capital: '曼谷',
        population: '7000万',
        area: '51.3万平方公里',
        highlight: '微笑之国，佛教文化与热带风情的完美结合',

        // 签证信息
        visa: {
            type: '落地签/电子签',
            duration: '15-30天',
            difficulty: '简单',
            requirements: [
                '护照原件（有效期6个月以上）',
                '2寸白底照片',
                '往返机票',
                '酒店预订单',
                '落地签费用2000泰铢'
            ],
            tips: '落地签方便快捷，建议提前准备好材料避免排队'
        },

        // 货币信息
        currency: {
            name: '泰铢',
            code: 'THB',
            symbol: '฿',
            rate: '1泰铢 ≈ 0.2人民币',
            paymentTips: '建议国内兑换或当地ATM取现，小商贩主要收现金，大商场可用信用卡'
        },

        // 语言信息
        language: {
            official: '泰语',
            commonGreetings: [
                { chinese: '你好', thai: 'สวัสดี (Sa-wat-dee)' },
                { chinese: '谢谢', thai: 'ขอบคุณ (Kob-kun)' },
                { chinese: '对不起', thai: 'ขอโทษ (Kor-tot)' }
            ],
            usefulPhrases: [
                { chinese: '多少钱？', thai: 'เท่าไหร่ (Tao-rai?)' },
                { chinese: '我不会说泰语', thai: 'ผม/ดิฉัน พูดไทยไม่ได้ (Pom/Di-chun phut thai mai dai)' }
            ]
        },

        // 习俗禁忌
        customs: [
            {
                title: '佛教文化',
                content: '进入寺庙要脱鞋，不要触摸和尚，女性不能触碰和尚身体'
            },
            {
                title: '皇室尊重',
                content: '不要议论皇室，货币上印有国王头像，要表示尊重'
            },
            {
                title: '社交礼仪',
                content: '见面行合十礼，不要摸别人的头，脚不能指向人或佛像'
            }
        ],

        // 玩法路线
        routes: [
            {
                title: '曼谷-普吉岛7日游',
                duration: '7天6夜',
                highlights: ['大皇宫', '四面佛', '普吉岛海滩', '皮皮岛', '泰式按摩'],
                description: '文化体验与海岛度假的完美结合'
            },
            {
                title: '清迈古城5日游',
                duration: '5天4夜',
                highlights: ['清迈古城', '素贴山', '夜间动物园', '长颈族村', '丛林飞跃'],
                description: '体验泰北文化，享受宁静的古城生活'
            }
        ],

        // 高频词汇
        vocabulary: [
            { chinese: '冬阴功汤', english: 'Tom Yum Goong' },
            { chinese: '泰式炒河粉', english: 'Pad Thai' },
            { chinese: '芒果糯米饭', english: 'Mango Sticky Rice' },
            { chinese: '按摩', english: 'Massage' },
            { chinese: '嘟嘟车', english: 'Tuk-tuk' },
            { chinese: '7-11便利店', english: '7-Eleven' }
        ]
    }
];

// Make data available globally
window.countriesData = countriesData;