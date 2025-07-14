/**
 * @fileoverview 統一モーダル管理システム - GASベストプラクティス準拠
 * プライバシー・デジタルシティズンシップモーダルの効率的管理
 */

/**
 * モーダル設定とコンテンツ管理
 */
const MODAL_CONFIG = {
  // モーダル種別定義
  TYPES: {
    PRIVACY_POLICY: 'privacy-policy',
    DIGITAL_CITIZENSHIP: 'digital-citizenship',
    TEACHER_GUIDE: 'teacher-guide',
    SECURITY_INFO: 'security-info',
    USAGE_GUIDE: 'usage-guide',
    CONFIRMATION: 'confirmation',
    ALERT: 'alert',
    LOADING: 'loading'
  },
  
  // 年齢層別設定
  AGE_GROUPS: {
    ELEMENTARY_LOWER: 'elementary-lower',    // 小学校低学年（1-3年）
    ELEMENTARY_UPPER: 'elementary-upper',    // 小学校高学年（4-6年） 
    JUNIOR_HIGH: 'junior-high',             // 中学校
    HIGH_SCHOOL: 'high-school',             // 高校
    TEACHER: 'teacher'                      // 教師向け
  },
  
  // 表示タイミング
  TRIGGER_TIMING: {
    FIRST_VISIT: 'first-visit',
    BEFORE_CRITICAL_ACTION: 'before-critical-action',
    MANUAL: 'manual',
    PERIODIC_REMINDER: 'periodic-reminder'
  },
  
  // キャッシュ設定
  CACHE_TTL: 300, // 5分間
  
  // アニメーション設定
  ANIMATION: {
    FADE_DURATION: 300,
    SLIDE_DURATION: 250
  }
};

/**
 * モーダルコンテンツテンプレート
 * 年齢層・用途別のコンテンツ定義
 */
const MODAL_CONTENT_TEMPLATES = {
  
  // プライバシーポリシー（年齢層別）
  [MODAL_CONFIG.TYPES.PRIVACY_POLICY]: {
    
    [MODAL_CONFIG.AGE_GROUPS.ELEMENTARY_LOWER]: {
      title: "🔒 あんしんして つかうために",
      content: `
        <div class="space-y-4 text-lg">
          <div class="bg-blue-100 dark:bg-blue-900/30 rounded-lg p-4">
            <h4 class="font-bold text-blue-600 dark:text-blue-400 mb-2">📝 なまえや こたえについて</h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li>• きみの なまえや こたえは、せんせいだけが みることができます</li>
              <li>• ほかの ひとには みせません</li>
              <li>• べんきょうのためだけに つかいます</li>
            </ul>
          </div>
          
          <div class="bg-green-100 dark:bg-green-900/30 rounded-lg p-4">
            <h4 class="font-bold text-green-600 dark:text-green-400 mb-2">🛡️ あんぜんな つかいかた</h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li>• じぶんの ひみつや おうちのことは かかないでね</li>
              <li>• ともだちの わるぐちは かかないでね</li>
              <li>• こまったら せんせいに そうだんしてね</li>
            </ul>
          </div>
          
          <div class="bg-yellow-100 dark:bg-yellow-900/30 rounded-lg p-4">
            <h4 class="font-bold text-yellow-600 dark:text-yellow-400 mb-2">💝 やくそく</h4>
            <p class="text-gray-700 dark:text-gray-300">
              みんなが たのしく べんきょうできるように、
              やさしい きもちで つかいましょう。
            </p>
          </div>
        </div>
      `,
      actionText: "わかったよ！",
      requiresParental: true
    },
    
    [MODAL_CONFIG.AGE_GROUPS.ELEMENTARY_UPPER]: {
      title: "🔒 プライバシーとデータ保護について",
      content: `
        <div class="space-y-4">
          <div class="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3 flex items-center">
              <svg class="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"></path>
              </svg>
              個人情報の取り扱い
            </h4>
            <ul class="space-y-2 text-sm text-gray-700 dark:text-gray-300">
              <li><strong>収集する情報</strong>：名前、クラス、学習への回答・意見</li>
              <li><strong>利用目的</strong>：学習活動の支援と教育効果の向上</li>
              <li><strong>保存期間</strong>：学年終了まで（その後は適切に削除）</li>
              <li><strong>共有範囲</strong>：担当の先生のみ（保護者には必要に応じて）</li>
            </ul>
          </div>
          
          <div class="bg-green-50 dark:bg-green-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3">🛡️ 安全な利用のために</h4>
            <ul class="space-y-2 text-sm text-gray-700 dark:text-gray-300">
              <li>• 住所、電話番号、家族の情報などは入力しないでください</li>
              <li>• 他の人を傷つける言葉や悪口は書かないでください</li>
              <li>• わからないことがあれば、必ず先生に相談してください</li>
              <li>• このシステムは学校の授業でのみ使用してください</li>
            </ul>
          </div>
          
          <div class="bg-purple-50 dark:bg-purple-900/20 rounded-lg p-4 border border-purple-200 dark:border-purple-800">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3">📚 学習での活用</h4>
            <p class="text-sm text-gray-700 dark:text-gray-300">
              この回答ボードは、みなさんの意見や考えを共有し、
              お互いから学び合うためのツールです。
              正解・不正解を気にせず、自分の考えを積極的に表現してください。
            </p>
          </div>
        </div>
      `,
      actionText: "理解しました",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.JUNIOR_HIGH]: {
      title: "🔒 プライバシーポリシーと個人情報保護",
      content: `
        <div class="space-y-4">
          <div class="border-l-4 border-blue-500 bg-blue-50 dark:bg-blue-900/20 p-4">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3">個人情報の取り扱いについて</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <p><strong>収集する情報：</strong></p>
              <ul class="ml-4 space-y-1">
                <li>• 氏名、学年・クラス</li>
                <li>• 学習活動での回答・意見・感想</li>
                <li>• システム利用ログ（アクセス日時等）</li>
              </ul>
              
              <p><strong>利用目的：</strong></p>
              <ul class="ml-4 space-y-1">
                <li>• 学習活動の円滑な実施と効果測定</li>
                <li>• 個別指導・評価への活用</li>
                <li>• システムの安全性確保</li>
              </ul>
              
              <p><strong>第三者提供：</strong></p>
              <ul class="ml-4 space-y-1">
                <li>• 原則として第三者には提供しません</li>
                <li>• 保護者・学校管理者には教育上必要な範囲で提供する場合があります</li>
                <li>• 法令に基づく場合を除きます</li>
              </ul>
            </div>
          </div>
          
          <div class="border-l-4 border-green-500 bg-green-50 dark:bg-green-900/20 p-4">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3">あなたの権利</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <ul class="space-y-1">
                <li>• <strong>アクセス権</strong>：自分の情報の確認を求めることができます</li>
                <li>• <strong>訂正権</strong>：間違った情報の訂正を求めることができます</li>
                <li>• <strong>削除権</strong>：不要になった情報の削除を求めることができます</li>
                <li>• <strong>利用停止権</strong>：情報の利用停止を求めることができます</li>
              </ul>
              <p class="mt-2 text-xs text-gray-600 dark:text-gray-400">
                ※これらの権利を行使したい場合は、担当の先生に相談してください
              </p>
            </div>
          </div>
          
          <div class="border-l-4 border-purple-500 bg-purple-50 dark:bg-purple-900/20 p-4">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3">データセキュリティ</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <ul class="space-y-1">
                <li>• すべての通信は暗号化されています（HTTPS）</li>
                <li>• Google Workspace for Educationのセキュリティ基準に準拠</li>
                <li>• 定期的なセキュリティ更新と監視を実施</li>
                <li>• アクセスログの記録と分析による不正防止</li>
              </ul>
            </div>
          </div>
        </div>
      `,
      actionText: "同意して続ける",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.HIGH_SCHOOL]: {
      title: "🔒 データ保護と技術的セキュリティについて",
      content: `
        <div class="space-y-4">
          <div class="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m5.652-1.652a4.012 4.012 0 000 5.304m-4.306-5.304a4.012 4.012 0 000 5.304M6.343 6.343a8.018 8.018 0 000 11.314m11.314-11.314a8.018 8.018 0 000 11.314"></path>
              </svg>
              技術的保護仕組みの理解
            </h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <p>このシステムがどのように皆さんのデータを技術的に保護しているかを理解しましょう：</p>
              <ul class="ml-4 space-y-1">
                <li>• <strong>暗号化通信（HTTPS/TLS）</strong>：すべてのデータ送受信は暗号化されています</li>
                <li>• <strong>アクセス制御</strong>：認証されたユーザーのみがデータにアクセス可能</li>
                <li>• <strong>データ分離</strong>：各クラス・学校のデータは完全に分離されています</li>
                <li>• <strong>自動バックアップ</strong>：データ消失を防ぐ定期的なバックアップ</li>
              </ul>
            </div>
          </div>
          
          <div class="bg-green-50 dark:bg-green-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path>
              </svg>
              データの教育的活用
            </h4>
            <div class="text-sm space-y-3">
              <div>
                <p class="font-semibold text-gray-800 dark:text-gray-200">学習支援のための情報活用</p>
                <ul class="ml-4 mt-1 space-y-1 text-gray-700 dark:text-gray-300">
                  <li>• <strong>学習履歴の可視化</strong>：自分の成長や理解度の把握</li>
                  <li>• <strong>協働学習の促進</strong>：クラスメートとの意見交換と相互学習</li>
                  <li>• <strong>個別指導の最適化</strong>：先生が効果的な指導を行うための分析</li>
                  <li>• <strong>学習環境の改善</strong>：システムの使いやすさとセキュリティの向上</li>
                </ul>
              </div>
              
              <div>
                <p class="font-semibold text-gray-800 dark:text-gray-200">技術的なデータ管理</p>
                <ul class="ml-4 mt-1 space-y-1 text-gray-700 dark:text-gray-300">
                  <li>• <strong>データの最小化</strong>：学習に必要な情報のみを収集・保存</li>
                  <li>• <strong>アクセス制限</strong>：担当教師と本人のみがデータにアクセス可能</li>
                  <li>• <strong>自動削除</strong>：不要になったデータは自動的に安全に削除</li>
                  <li>• <strong>クラウド保護</strong>：Google の高度なセキュリティ技術で保護</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="bg-purple-50 dark:bg-purple-900/20 rounded-lg p-4 border border-purple-200 dark:border-purple-800">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197m13.5-9a2.5 2.5 0 11-5 0 2.5 2.5 0 015 0z"></path>
              </svg>
              自分のデータをコントロールする権利
            </h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-3">
              <div>
                <p class="font-semibold text-purple-600 dark:text-purple-400">学習者としてできること</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• <strong>データの確認</strong>：自分の学習データや投稿履歴を確認</li>
                  <li>• <strong>間違いの修正</strong>：誤った情報の訂正を先生に依頼</li>
                  <li>• <strong>削除の要求</strong>：不要になったデータの削除を依頼</li>
                  <li>• <strong>学習データの取得</strong>：自分の成長記録をファイルで受け取り</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold text-purple-600 dark:text-purple-400">相談・質問窓口</p>
                <div class="bg-purple-100 dark:bg-purple-800 p-3 rounded mt-2">
                  <p class="text-sm">
                    <strong>まずは担当の先生に相談してください</strong><br>
                    分からないことや心配なことがあれば、いつでも気軽に質問できます。
                    先生が解決できない技術的な問題は、システム管理者と連携して対応します。
                  </p>
                </div>
              </div>
            </div>
          </div>
          
          <div class="bg-orange-50 dark:bg-orange-900/20 rounded-lg p-4 border border-orange-200 dark:border-orange-800">
            <h4 class="font-bold text-orange-700 dark:text-orange-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path>
              </svg>
              将来への準備：デジタル社会でのデータ管理
            </h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <p>このシステムを通じて、将来のデジタル社会で必要な「データ管理意識」を身につけましょう：</p>
              <ul class="ml-4 mt-2 space-y-1">
                <li>• <strong>データの価値理解</strong>：個人情報がなぜ大切なのかを実体験で学習</li>
                <li>• <strong>適切な共有判断</strong>：何を共有し、何を秘密にするかの判断力を養成</li>
                <li>• <strong>セキュリティ意識</strong>：安全な技術の仕組みを理解し活用する能力</li>
                <li>• <strong>デジタル・エチケット</strong>：オンライン環境での礼儀と責任ある行動</li>
              </ul>
            </div>
          </div>
        </div>
      `,
      actionText: "同意して利用開始",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.TEACHER]: {
      title: "🔒 教師向け：プライバシーポリシーと管理責任",
      content: `
        <div class="space-y-4">
          <div class="bg-red-50 dark:bg-red-900/20 rounded-lg p-4 border border-red-200 dark:border-red-800">
            <h4 class="font-bold text-red-700 dark:text-red-300 mb-3">⚠️ 管理者としての責任</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <ul class="space-y-1">
                <li>• 生徒の個人情報を適切に保護し、目的外利用を防ぐ責任があります</li>
                <li>• システムの設定・公開範囲を適切に管理してください</li>
                <li>• 不適切な投稿を発見した場合は速やかに対処してください</li>
                <li>• 保護者への説明と同意取得が必要な場合があります</li>
              </ul>
            </div>
          </div>
          
          <div class="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3">📊 データ管理のガイドライン</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300 space-y-2">
              <ul class="space-y-1">
                <li>• 収集したデータは教育目的でのみ利用してください</li>
                <li>• 第三者への提供は学校の規定に従って判断してください</li>
                <li>• 定期的にデータの整理・削除を行ってください</li>
                <li>• セキュリティインシデント発生時は速やかに報告してください</li>
              </ul>
            </div>
          </div>
          
          <div class="bg-green-50 dark:bg-green-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3">🛡️ セキュリティ設定の確認</h4>
            <div class="text-sm text-gray-700 dark:text-gray-300">
              <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
                <div>
                  <p class="font-semibold mb-1">必須設定</p>
                  <ul class="space-y-1 text-xs">
                    <li>✓ スプレッドシート共有範囲の確認</li>
                    <li>✓ 回答ボード公開範囲の設定</li>
                    <li>✓ 名前表示・非表示の選択</li>
                  </ul>
                </div>
                <div>
                  <p class="font-semibold mb-1">推奨設定</p>
                  <ul class="space-y-1 text-xs">
                    <li>✓ 定期的なデータバックアップ</li>
                    <li>✓ アクセスログの確認</li>
                    <li>✓ 生徒向けガイダンスの実施</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      `,
      actionText: "理解しました（管理開始）",
      requiresParental: false
    }
  },
  
  // デジタル・シティズンシップ（年齢層別）
  [MODAL_CONFIG.TYPES.DIGITAL_CITIZENSHIP]: {
    
    [MODAL_CONFIG.AGE_GROUPS.ELEMENTARY_LOWER]: {
      title: "🌐 デジタルせかいの やくそく",
      content: `
        <div class="space-y-4">
          <div class="bg-blue-100 dark:bg-blue-900/30 rounded-lg p-4">
            <h4 class="font-bold text-blue-600 dark:text-blue-400 mb-2 text-xl">🤝 やさしい きもち</h4>
            <div class="text-lg text-gray-700 dark:text-gray-300 space-y-2">
              <p>• ともだちに やさしい ことばを つかおう</p>
              <p>• みんなの かんがえを きこう</p>
              <p>• ちがう いけんも だいじにしよう</p>
            </div>
          </div>
          
          <div class="bg-green-100 dark:bg-green-900/30 rounded-lg p-4">
            <h4 class="font-bold text-green-600 dark:text-green-400 mb-2 text-xl">🛡️ あんぜんに つかう</h4>
            <div class="text-lg text-gray-700 dark:text-gray-300 space-y-2">
              <p>• ひみつの ことは かかない</p>
              <p>• わるい ことばは つかわない</p>
              <p>• こまったら せんせいに いう</p>
            </div>
          </div>
          
          <div class="bg-yellow-100 dark:bg-yellow-900/30 rounded-lg p-4">
            <h4 class="font-bold text-yellow-600 dark:text-yellow-400 mb-2 text-xl">📚 たくさん まなぼう</h4>
            <div class="text-lg text-gray-700 dark:text-gray-300 space-y-2">
              <p>• たくさん かんがえよう</p>
              <p>• じぶんの いけんを いおう</p>
              <p>• ともだちから まなぼう</p>
            </div>
          </div>
          
          <div class="text-center bg-purple-100 dark:bg-purple-900/30 rounded-lg p-4">
            <p class="text-xl text-purple-700 dark:text-purple-300 font-bold">
              みんなで たのしく べんきょうしよう！ 🌟
            </p>
          </div>
        </div>
      `,
      actionText: "やくそくするよ！",
      requiresParental: true
    },
    
    [MODAL_CONFIG.AGE_GROUPS.ELEMENTARY_UPPER]: {
      title: "🌐 デジタル・シティズンシップを身につけよう",
      content: `
        <div class="space-y-4">
          <div class="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3 flex items-center">
              <span class="text-2xl mr-2">🤝</span>
              デジタル・コミュニケーション
            </h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li><strong>相手を思いやる：</strong>画面の向こうには本当の人がいることを忘れずに</li>
              <li><strong>建設的な対話：</strong>違う意見も尊重し、みんなで学び合う姿勢を大切に</li>
              <li><strong>適切な言葉遣い：</strong>友達との普段の会話と同じように、丁寧で優しい言葉を使う</li>
            </ul>
          </div>
          
          <div class="bg-green-50 dark:bg-green-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3 flex items-center">
              <span class="text-2xl mr-2">🛡️</span>
              デジタル・セキュリティ
            </h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li><strong>個人情報の保護：</strong>住所、電話番号、家族の情報は書かない</li>
              <li><strong>安全な共有：</strong>学習に関することだけを投稿する</li>
              <li><strong>困った時の対応：</strong>わからないことがあれば必ず先生に相談</li>
            </ul>
          </div>
          
          <div class="bg-purple-50 dark:bg-purple-900/20 rounded-lg p-4 border border-purple-200 dark:border-purple-800">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3 flex items-center">
              <span class="text-2xl mr-2">📚</span>
              デジタル・リテラシー
            </h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li><strong>情報の確認：</strong>何かを書く前に、本当のことかどうか考える</li>
              <li><strong>批判的思考：</strong>他の人の意見も聞いて、自分なりに考える</li>
              <li><strong>建設的な議論：</strong>根拠を持って、わかりやすく自分の考えを伝える</li>
            </ul>
          </div>
          
          <div class="bg-yellow-50 dark:bg-yellow-900/20 rounded-lg p-4 border border-yellow-200 dark:border-yellow-800">
            <h4 class="font-bold text-yellow-700 dark:text-yellow-300 mb-3 flex items-center">
              <span class="text-2xl mr-2">⚖️</span>
              デジタル・責任
            </h4>
            <ul class="space-y-2 text-gray-700 dark:text-gray-300">
              <li><strong>自分の行動に責任を持つ：</strong>デジタル空間でも現実と同じルールがある</li>
              <li><strong>記録が残ることを理解：</strong>書いたことは記録に残ることを覚えておく</li>
              <li><strong>良い学習環境を作る：</strong>みんなが気持ちよく学習できる環境作りに協力</li>
            </ul>
          </div>
        </div>
      `,
      actionText: "これらを守って参加します",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.JUNIOR_HIGH]: {
      title: "🌐 デジタル・シティズンシップの実践",
      content: `
        <div class="space-y-4">
          <div class="border-l-4 border-blue-500 bg-blue-50 dark:bg-blue-900/20 p-4">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3">💬 建設的なコミュニケーション</h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold">多様性の尊重</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 異なる意見や立場を理解しようとする姿勢</li>
                  <li>• 文化的・個人的背景の違いへの配慮</li>
                  <li>• 建設的な批判と人格攻撃の区別</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold">効果的な表現</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 根拠を示した論理的な意見表明</li>
                  <li>• 相手にわかりやすい言葉の選択</li>
                  <li>• 感情的にならず冷静な議論</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="border-l-4 border-green-500 bg-green-50 dark:bg-green-900/20 p-4">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3">🔍 情報リテラシー</h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold">情報の評価と検証</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 情報源の信頼性確認</li>
                  <li>• 複数の視点からの情報収集</li>
                  <li>• 偏見や先入観への注意</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold">適切な情報共有</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 著作権や肖像権の尊重</li>
                  <li>• 不確実な情報の拡散防止</li>
                  <li>• プライバシーへの配慮</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="border-l-4 border-purple-500 bg-purple-50 dark:bg-purple-900/20 p-4">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3">🤝 デジタル倫理と社会的責任</h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold">倫理的な判断力の育成</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• <strong>他者への思いやり</strong>：相手の気持ちを考えた言葉選び</li>
                  <li>• <strong>公正性の重視</strong>：偏見のない客観的な意見表明</li>
                  <li>• <strong>誠実性の実践</strong>：嘘や誇張のない正直なコミュニケーション</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold">社会的責任の意識</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• <strong>コミュニティへの貢献</strong>：みんなの学習を支援する行動</li>
                  <li>• <strong>建設的な議論</strong>：問題解決につながる前向きな対話</li>
                  <li>• <strong>継続的な成長</strong>：失敗から学び、改善し続ける姿勢</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="border-l-4 border-orange-500 bg-orange-50 dark:bg-orange-900/20 p-4">
            <h4 class="font-bold text-orange-700 dark:text-orange-300 mb-3">🚀 創造的な活用</h4>
            <div class="space-y-2 text-sm text-gray-700 dark:text-gray-300">
              <ul class="space-y-1">
                <li>• デジタル技術を活用した学習の深化</li>
                <li>• 協働学習でのリーダーシップ発揮</li>
                <li>• 問題解決のためのツール活用</li>
                <li>• 創造的な表現と発信</li>
              </ul>
            </div>
          </div>
        </div>
      `,
      actionText: "実践していきます",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.HIGH_SCHOOL]: {
      title: "🌐 デジタル・シティズンシップ：社会参画への準備",
      content: `
        <div class="space-y-4">
          <div class="bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-blue-900/20 dark:to-indigo-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z"></path>
              </svg>
              グローバル・デジタル・コミュニティでの責任
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-indigo-700 dark:text-indigo-300">文化横断的コミュニケーション</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 異文化間での価値観の違いを理解し、尊重する姿勢</li>
                  <li>• 言語や表現方法の多様性への配慮と適応</li>
                  <li>• ステレオタイプや偏見を排除した客観的な議論</li>
                  <li>• グローバルな視点と地域的な視点のバランス</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold text-indigo-700 dark:text-indigo-300">社会課題への参画</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• デジタル格差の理解と解決への意識</li>
                  <li>• 持続可能な開発目標（SDGs）との関連づけ</li>
                  <li>• 情報アクセシビリティへの配慮</li>
                  <li>• 社会変革のためのテクノロジー活用</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="bg-gradient-to-r from-green-50 to-emerald-50 dark:from-green-900/20 dark:to-emerald-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4M7.835 4.697a3.42 3.42 0 001.946-.806 3.42 3.42 0 014.438 0 3.42 3.42 0 001.946.806 3.42 3.42 0 013.138 3.138 3.42 3.42 0 00.806 1.946 3.42 3.42 0 010 4.438 3.42 3.42 0 00-.806 1.946 3.42 3.42 0 01-3.138 3.138 3.42 3.42 0 00-1.946.806 3.42 3.42 0 01-4.438 0 3.42 3.42 0 00-1.946-.806 3.42 3.42 0 01-3.138-3.138 3.42 3.42 0 00-.806-1.946 3.42 3.42 0 010-4.438 3.42 3.42 0 00.806-1.946 3.42 3.42 0 013.138-3.138z"></path>
              </svg>
              高度な情報リテラシーと批判的思考
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-emerald-700 dark:text-emerald-300">メディア・リテラシー</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• フェイクニュースや偽情報の識別と対策</li>
                  <li>• エコーチェンバー効果の理解と回避</li>
                  <li>• アルゴリズムバイアスの認識と対処</li>
                  <li>• 複数の情報源からの総合的判断</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold text-emerald-700 dark:text-emerald-300">デジタル・リーダーシップ</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• オンライン学習コミュニティの活性化</li>
                  <li>• 建設的なディスカッションの促進</li>
                  <li>• 知識共有と協働学習の推進</li>
                  <li>• 他者の学習支援とメンターシップ</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-900/20 dark:to-pink-900/20 rounded-lg p-4 border border-purple-200 dark:border-purple-800">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.246 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"></path>
              </svg>
              将来への準備：デジタル社会でのキャリア設計
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-pink-700 dark:text-pink-300">デジタル・スキルセット</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 情報の収集・分析・活用能力の向上</li>
                  <li>• デジタル・コンテンツ制作スキル</li>
                  <li>• オンライン・プレゼンテーション能力</li>
                  <li>• ネットワーキングとコラボレーション</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold text-pink-700 dark:text-pink-300">職業倫理の形成</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 知的財産権の理解と尊重</li>
                  <li>• 企業秘密・機密情報の取り扱い</li>
                  <li>• プロフェッショナルなオンライン・プレゼンス</li>
                  <li>• 継続的学習とスキルアップデート</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="bg-gradient-to-r from-orange-50 to-red-50 dark:from-orange-900/20 dark:to-red-900/20 rounded-lg p-4 border border-orange-200 dark:border-orange-800">
            <h4 class="font-bold text-orange-700 dark:text-orange-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path>
              </svg>
              実践チャレンジ：今日から始めるデジタル・シティズンシップ
            </h4>
            <div class="text-sm text-gray-700 dark:text-gray-300">
              <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div class="bg-white dark:bg-gray-800 rounded p-3">
                  <p class="font-semibold text-orange-600 dark:text-orange-400 mb-2">今日のアクション</p>
                  <ul class="space-y-1 text-xs">
                    <li>✓ 建設的なフィードバックを1つ投稿</li>
                    <li>✓ 異なる視点を持つ意見を尊重</li>
                    <li>✓ 情報源を確認してから発言</li>
                  </ul>
                </div>
                <div class="bg-white dark:bg-gray-800 rounded p-3">
                  <p class="font-semibold text-red-600 dark:text-red-400 mb-2">今週の目標</p>
                  <ul class="space-y-1 text-xs">
                    <li>✓ 1つの社会課題について深く調べる</li>
                    <li>✓ 他者の学習支援を積極的に行う</li>
                    <li>✓ デジタル・フットプリントを意識</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      `,
      actionText: "デジタル・リーダーとして参画",
      requiresParental: false
    },
    
    [MODAL_CONFIG.AGE_GROUPS.TEACHER]: {
      title: "🌐 教師向け：デジタル・シティズンシップ指導ガイド",
      content: `
        <div class="space-y-4">
          <div class="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 border border-blue-200 dark:border-blue-800">
            <h4 class="font-bold text-blue-700 dark:text-blue-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-4m-5 0H3m2 0h4M9 7h6m-6 4h6m-2 4h2M9 15h2"></path>
              </svg>
              指導フレームワーク
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-blue-600 dark:text-blue-400">段階的指導アプローチ</p>
                <div class="ml-4 mt-2 space-y-2">
                  <div class="border-l-2 border-blue-300 pl-3">
                    <p class="font-medium">Phase 1: 意識化</p>
                    <p class="text-xs">デジタル行動の影響を理解させる</p>
                  </div>
                  <div class="border-l-2 border-green-300 pl-3">
                    <p class="font-medium">Phase 2: スキル習得</p>
                    <p class="text-xs">具体的なデジタル・スキルを身につける</p>
                  </div>
                  <div class="border-l-2 border-purple-300 pl-3">
                    <p class="font-medium">Phase 3: 実践応用</p>
                    <p class="text-xs">学習活動での継続的な実践</p>
                  </div>
                  <div class="border-l-2 border-orange-300 pl-3">
                    <p class="font-medium">Phase 4: リーダーシップ</p>
                    <p class="text-xs">他者への指導・支援役割の発揮</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
          <div class="bg-green-50 dark:bg-green-900/20 rounded-lg p-4 border border-green-200 dark:border-green-800">
            <h4 class="font-bold text-green-700 dark:text-green-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"></path>
              </svg>
              実践的指導テクニック
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-green-600 dark:text-green-400">リアルタイム指導</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 不適切な投稿を発見した際の即座のフィードバック</li>
                  <li>• 建設的な議論を促すファシリテーション</li>
                  <li>• 学習者同士の相互支援を促進する声かけ</li>
                  <li>• ポジティブな行動の積極的な認知と称賛</li>
                </ul>
              </div>
              <div>
                <p class="font-semibold text-green-600 dark:text-green-400">事例ベース学習</p>
                <ul class="ml-4 mt-1 space-y-1">
                  <li>• 実際の投稿を使った良い例・悪い例の分析</li>
                  <li>• 炎上事例の検証と防止策の検討</li>
                  <li>• ロールプレイングによる立場の違いの体験</li>
                  <li>• 成功事例の共有と要因分析</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="bg-purple-50 dark:bg-purple-900/20 rounded-lg p-4 border border-purple-200 dark:border-purple-800">
            <h4 class="font-bold text-purple-700 dark:text-purple-300 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 4a2 2 0 114 0v1a1 1 0 001 1h3a1 1 0 011 1v3a1 1 0 01-1 1h-1a2 2 0 100 4h1a1 1 0 011 1v3a1 1 0 01-1 1h-3a1 1 0 01-1-1v-1a2 2 0 10-4 0v1a1 1 0 01-1 1H7a1 1 0 01-1-1v-3a1 1 0 011-1h1a2 2 0 100-4H7a1 1 0 01-1-1V7a1 1 0 011-1h3a1 1 0 001-1V4z"></path>
              </svg>
              課題対応ガイドライン
            </h4>
            <div class="space-y-3 text-sm text-gray-700 dark:text-gray-300">
              <div>
                <p class="font-semibold text-purple-600 dark:text-purple-400">レベル別対応方針</p>
                <div class="space-y-2 mt-2">
                  <div class="bg-yellow-100 dark:bg-yellow-900/30 p-2 rounded">
                    <p class="font-medium text-yellow-800 dark:text-yellow-300">軽微（予防的指導）</p>
                    <p class="text-xs">個別の声かけ、正しい方法の教示</p>
                  </div>
                  <div class="bg-orange-100 dark:bg-orange-900/30 p-2 rounded">
                    <p class="font-medium text-orange-800 dark:text-orange-300">中程度（教育的介入）</p>
                    <p class="text-xs">全体指導、事例研究、保護者連絡</p>
                  </div>
                  <div class="bg-red-100 dark:bg-red-900/30 p-2 rounded">
                    <p class="font-medium text-red-800 dark:text-red-300">深刻（組織的対応）</p>
                    <p class="text-xs">管理職報告、専門機関連携、システム利用停止</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
          <div class="bg-gray-50 dark:bg-gray-800 rounded-lg p-4 border">
            <h4 class="font-bold text-gray-800 dark:text-gray-200 mb-3 flex items-center">
              <svg class="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path>
              </svg>
              継続的改善のためのチェックリスト
            </h4>
            <div class="space-y-2 text-sm text-gray-700 dark:text-gray-300">
              <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div>
                  <p class="font-semibold mb-2">日次確認項目</p>
                  <ul class="space-y-1 text-xs">
                    <li>□ 投稿内容の適切性確認</li>
                    <li>□ 学習者間の相互作用観察</li>
                    <li>□ ポジティブな行動の記録</li>
                    <li>□ 改善が必要な点の特定</li>
                  </ul>
                </div>
                <div>
                  <p class="font-semibold mb-2">週次振り返り</p>
                  <ul class="space-y-1 text-xs">
                    <li>□ 学習者の成長度合い評価</li>
                    <li>□ 指導方法の効果検証</li>
                    <li>□ 保護者との情報共有</li>
                    <li>□ 次週の指導計画立案</li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        </div>
      `,
      actionText: "指導実践を開始",
      requiresParental: false
    }
  }
};

/**
 * 統一モーダル管理クラス
 * GASベストプラクティスに準拠した効率的な実装
 */
class UnifiedModalManager {
  constructor() {
    this.modals = new Map();
    this.eventListeners = new Map();
    this.config = MODAL_CONFIG;
    this.contentTemplates = MODAL_CONTENT_TEMPLATES;
    this.currentAge = MODAL_CONFIG.AGE_GROUPS.JUNIOR_HIGH; // デフォルト
    this.isInitialized = false;
    
    // パフォーマンス監視
    this.performanceMetrics = {
      showCount: 0,
      hideCount: 0,
      errorCount: 0,
      avgLoadTime: 0
    };
  }

  /**
   * モーダルシステムの初期化
   * @param {object} options - 初期化オプション
   */
  initialize(options = {}) {
    if (this.isInitialized) {
      console.warn('[ModalManager] Already initialized');
      return;
    }

    try {
      const startTime = performance.now();
      
      // 設定の適用
      this.currentAge = options.ageGroup || this.currentAge;
      
      // DOM要素のキャッシュ
      this.cacheModalElements();
      
      // グローバルイベントリスナーの設定
      this.setupGlobalEventListeners();
      
      // パフォーマンス測定
      const loadTime = performance.now() - startTime;
      this.performanceMetrics.avgLoadTime = loadTime;
      
      this.isInitialized = true;
      console.log(`[ModalManager] Initialized successfully in ${loadTime.toFixed(2)}ms`);
      
    } catch (error) {
      this.performanceMetrics.errorCount++;
      console.error('[ModalManager] Initialization failed:', error);
      throw new Error(`モーダルシステムの初期化に失敗しました: ${error.message}`);
    }
  }

  /**
   * プライバシーモーダルの表示
   * @param {object} options - 表示オプション
   */
  showPrivacyModal(options = {}) {
    const modalOptions = {
      type: this.config.TYPES.PRIVACY_POLICY,
      ageGroup: options.ageGroup || this.currentAge,
      onContinue: options.onContinue,
      onCancel: options.onCancel,
      ...options
    };
    
    return this.showModal(modalOptions);
  }

  /**
   * デジタルシティズンシップモーダルの表示
   * @param {object} options - 表示オプション
   */
  showDigitalCitizenshipModal(options = {}) {
    const modalOptions = {
      type: this.config.TYPES.DIGITAL_CITIZENSHIP,
      ageGroup: options.ageGroup || this.currentAge,
      onAccept: options.onAccept,
      onCancel: options.onCancel,
      ...options
    };
    
    return this.showModal(modalOptions);
  }

  /**
   * 統一モーダル表示メソッド - GASベストプラクティス準拠
   * @param {object} options - モーダル表示オプション
   * @returns {Promise} モーダル表示Promise
   */
  async showModal(options) {
    const operationId = this.generateOperationId();
    
    try {
      const startTime = Date.now(); // GAS環境でより安定
      this.performanceMetrics.showCount++;
      
      console.log(`[ModalManager:${operationId}] Starting modal display`, options);
      
      // オプションの検証（同期的に実行）
      this.validateModalOptions(options);
      
      // ロック機能付きコンテンツ取得（重複呼び出し防止）
      const content = await this.getModalContentWithErrorHandling(options.type, options.ageGroup, operationId);
      
      // 段階的モーダル構築（メモリ効率重視）
      const modalElement = await this.createModalElementSafely(options.type, content, options, operationId);
      
      // 非同期表示とイベント設定
      await this.displayModalWithErrorHandling(modalElement, options, operationId);
      
      // 成功メトリクスの記録
      const loadTime = Date.now() - startTime;
      this.updatePerformanceMetrics('show_success', loadTime);
      console.log(`[ModalManager:${operationId}] Modal displayed successfully in ${loadTime}ms`);
      
      // Promise の適切な処理
      return this.createModalPromiseWithTimeout(modalElement, options, operationId);
      
    } catch (error) {
      this.handleModalError(error, operationId, options);
      throw this.createUserFriendlyError(error, 'モーダル表示');
    }
  }

  /**
   * 操作IDの生成（デバッグとトレース用）
   * @returns {string} 一意の操作ID
   */
  generateOperationId() {
    return `modal_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * エラーハンドリング付きコンテンツ取得
   * @param {string} type - モーダルタイプ
   * @param {string} ageGroup - 年齢グループ
   * @param {string} operationId - 操作ID
   * @returns {Promise<Object>} コンテンツオブジェクト
   */
  async getModalContentWithErrorHandling(type, ageGroup, operationId) {
    try {
      const content = this.getModalContent(type, ageGroup);
      if (!content) {
        console.warn(`[ModalManager:${operationId}] No content found, using fallback`);
        return this.getFallbackContentSafe(type);
      }
      return content;
    } catch (error) {
      console.error(`[ModalManager:${operationId}] Content fetch error:`, error);
      return this.getFallbackContentSafe(type);
    }
  }

  /**
   * 安全なモーダル要素作成
   * @param {string} type - モーダルタイプ
   * @param {Object} content - コンテンツオブジェクト
   * @param {Object} options - オプション
   * @param {string} operationId - 操作ID
   * @returns {Promise<Element>} モーダル要素
   */
  async createModalElementSafely(type, content, options, operationId) {
    try {
      // 既存要素の確認とクリーンアップ
      const existingModal = document.getElementById(`unified-modal-${type}`);
      if (existingModal) {
        this.cleanupModalEventListeners(existingModal);
        existingModal.remove();
      }

      // 新しいモーダル要素の作成
      const modalElement = this.buildModalElement(type, content, options);
      modalElement.id = `unified-modal-${type}`;
      modalElement.setAttribute('data-operation-id', operationId);
      
      // DOM に追加
      document.body.appendChild(modalElement);
      
      console.log(`[ModalManager:${operationId}] Modal element created successfully`);
      return modalElement;
      
    } catch (error) {
      console.error(`[ModalManager:${operationId}] Modal element creation failed:`, error);
      throw new Error(`モーダル要素の作成に失敗しました: ${error.message}`);
    }
  }

  /**
   * エラーハンドリング付きモーダル表示
   * @param {Element} modalElement - モーダル要素
   * @param {Object} options - オプション
   * @param {string} operationId - 操作ID
   * @returns {Promise<void>}
   */
  async displayModalWithErrorHandling(modalElement, options, operationId) {
    try {
      // アニメーション前の準備
      modalElement.style.display = 'flex';
      modalElement.style.opacity = '0';
      
      // DOM 更新の確実な実行
      await this.nextTick();
      
      // フェードインアニメーション
      modalElement.style.transition = `opacity ${this.config.ANIMATION.FADE_DURATION}ms ease`;
      modalElement.style.opacity = '1';
      
      // アクセシビリティ設定
      this.setupAccessibility(modalElement);
      
      console.log(`[ModalManager:${operationId}] Modal displayed with animation`);
      
    } catch (error) {
      console.error(`[ModalManager:${operationId}] Modal display error:`, error);
      // フォールバック: アニメーションなしで表示
      modalElement.style.display = 'flex';
      modalElement.style.opacity = '1';
    }
  }

  /**
   * 次のイベントループまで待機
   * @returns {Promise<void>}
   */
  nextTick() {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  /**
   * モーダルコンテンツの取得
   * @param {string} type - モーダルタイプ
   * @param {string} ageGroup - 年齢層
   * @returns {object} コンテンツオブジェクト
   */
  getModalContent(type, ageGroup) {
    const typeContent = this.contentTemplates[type];
    if (!typeContent) {
      console.warn(`[ModalManager] Content type not found: ${type}`);
      return null;
    }
    
    const ageContent = typeContent[ageGroup];
    if (!ageContent) {
      console.warn(`[ModalManager] Age group content not found: ${type} - ${ageGroup}`);
      // フォールバック: 一般向けコンテンツを探す
      return typeContent[MODAL_CONFIG.AGE_GROUPS.JUNIOR_HIGH] || 
             typeContent[Object.keys(typeContent)[0]];
    }
    
    return ageContent;
  }

  /**
   * モーダル要素の作成または取得
   * @param {string} type - モーダルタイプ  
   * @param {object} content - コンテンツ
   * @param {object} options - オプション
   * @returns {HTMLElement} モーダル要素
   */
  async createOrGetModalElement(type, content, options) {
    const modalId = `unified-modal-${type}`;
    let modalElement = document.getElementById(modalId);
    
    if (!modalElement) {
      modalElement = this.createModalElement(modalId, content, options);
      document.body.appendChild(modalElement);
    } else {
      // 既存要素のコンテンツ更新
      this.updateModalContent(modalElement, content, options);
    }
    
    return modalElement;
  }

  /**
   * モーダル要素の作成
   * @param {string} modalId - モーダルID
   * @param {object} content - コンテンツ
   * @param {object} options - オプション
   * @returns {HTMLElement} 作成されたモーダル要素
   */
  createModalElement(modalId, content, options) {
    const modalElement = document.createElement('div');
    modalElement.id = modalId;
    modalElement.className = 'fixed inset-0 bg-black/80 z-[60] hidden';
    modalElement.setAttribute('role', 'dialog');
    modalElement.setAttribute('aria-modal', 'true');
    modalElement.setAttribute('aria-labelledby', `${modalId}-title`);
    
    modalElement.innerHTML = `
      <div class="flex items-center justify-center min-h-full p-4">
        <div class="glass-panel rounded-xl p-6 max-w-4xl w-full max-h-[90vh] overflow-y-auto">
          <div class="flex items-center justify-between mb-6">
            <h3 id="${modalId}-title" class="text-2xl font-bold text-cyan-400">
              ${this.escapeHtml(content.title)}
            </h3>
            <button type="button" class="modal-close-btn p-2 text-gray-400 transition-colors rounded-full hover:bg-gray-700 hover:text-white focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-cyan-500" aria-label="閉じる">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
              </svg>
            </button>
          </div>
          
          <div class="modal-content text-gray-300">
            ${content.content}
          </div>
          
          <div class="flex justify-end gap-3 mt-6">
            <button type="button" class="modal-cancel-btn btn btn-secondary">キャンセル</button>
            <button type="button" class="modal-action-btn btn btn-primary">
              ${this.escapeHtml(content.actionText)}
            </button>
          </div>
        </div>
      </div>
    `;
    
    return modalElement;
  }

  /**
   * モーダルの表示
   * @param {HTMLElement} modalElement - モーダル要素
   * @param {object} options - オプション
   */
  async displayModal(modalElement, options) {
    // アニメーション付きで表示
    modalElement.classList.remove('hidden');
    modalElement.style.opacity = '0';
    modalElement.style.transform = 'scale(0.95)';
    
    // フォーカス管理
    const firstFocusable = modalElement.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
    if (firstFocusable) {
      firstFocusable.focus();
    }
    
    // アニメーション実行
    return new Promise(resolve => {
      requestAnimationFrame(() => {
        modalElement.style.transition = `opacity ${this.config.ANIMATION.FADE_DURATION}ms ease, transform ${this.config.ANIMATION.SLIDE_DURATION}ms ease`;
        modalElement.style.opacity = '1';
        modalElement.style.transform = 'scale(1)';
        
        setTimeout(() => {
          modalElement.style.transition = '';
          resolve();
        }, Math.max(this.config.ANIMATION.FADE_DURATION, this.config.ANIMATION.SLIDE_DURATION));
      });
    });
  }

  /**
   * モーダル結果ハンドラーの設定
   * @param {HTMLElement} modalElement - モーダル要素
   * @param {object} options - オプション
   * @param {function} resolve - Promise resolve
   * @param {function} reject - Promise reject
   */
  setupModalResultHandlers(modalElement, options, resolve, reject) {
    const cleanup = () => {
      this.cleanupModalEventListeners(modalElement);
    };
    
    // アクションボタン
    const actionBtn = modalElement.querySelector('.modal-action-btn');
    if (actionBtn) {
      actionBtn.addEventListener('click', () => {
        cleanup();
        this.hideModal(modalElement);
        if (options.onContinue) options.onContinue();
        if (options.onAccept) options.onAccept();
        resolve({ result: 'accept', action: 'continue' });
      });
    }
    
    // キャンセルボタン
    const cancelBtn = modalElement.querySelector('.modal-cancel-btn');
    if (cancelBtn) {
      cancelBtn.addEventListener('click', () => {
        cleanup();
        this.hideModal(modalElement);
        if (options.onCancel) options.onCancel();
        resolve({ result: 'cancel', action: 'cancel' });
      });
    }
    
    // クローズボタン
    const closeBtn = modalElement.querySelector('.modal-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        cleanup();
        this.hideModal(modalElement);
        if (options.onCancel) options.onCancel();
        resolve({ result: 'close', action: 'close' });
      });
    }
    
    // ESCキー
    const escapeHandler = (e) => {
      if (e.key === 'Escape') {
        cleanup();
        this.hideModal(modalElement);
        if (options.onCancel) options.onCancel();
        resolve({ result: 'escape', action: 'escape' });
      }
    };
    document.addEventListener('keydown', escapeHandler);
    
    // クリーンアップ関数の追加
    const originalCleanup = cleanup;
    cleanup = () => {
      originalCleanup();
      document.removeEventListener('keydown', escapeHandler);
    };
  }

  /**
   * モーダルの非表示
   * @param {HTMLElement} modalElement - モーダル要素
   */
  async hideModal(modalElement) {
    this.performanceMetrics.hideCount++;
    
    return new Promise(resolve => {
      modalElement.style.transition = `opacity ${this.config.ANIMATION.FADE_DURATION}ms ease, transform ${this.config.ANIMATION.SLIDE_DURATION}ms ease`;
      modalElement.style.opacity = '0';
      modalElement.style.transform = 'scale(0.95)';
      
      setTimeout(() => {
        modalElement.classList.add('hidden');
        modalElement.style.transition = '';
        resolve();
      }, Math.max(this.config.ANIMATION.FADE_DURATION, this.config.ANIMATION.SLIDE_DURATION));
    });
  }

  /**
   * DOM要素のキャッシュ
   * @private
   */
  cacheModalElements() {
    // 必要に応じて要素をキャッシュ
    this.cachedElements = {
      body: document.body,
      focusableSelector: 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    };
  }

  /**
   * グローバルイベントリスナーの設定
   * @private
   */
  setupGlobalEventListeners() {
    // 必要最小限のグローバルリスナーを設定
    // 個別のモーダルイベントは動的に追加/削除
  }

  /**
   * モーダルオプションの検証
   * @private
   */
  validateModalOptions(options) {
    if (!options.type) {
      throw new Error('Modal type is required');
    }
    
    if (!this.contentTemplates[options.type]) {
      throw new Error(`Unknown modal type: ${options.type}`);
    }
  }

  /**
   * HTMLエスケープ
   * @private
   */
  escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * モーダルイベントリスナーのクリーンアップ
   * @private
   */
  cleanupModalEventListeners(modalElement) {
    // イベントリスナーを適切にクリーンアップ
    const buttons = modalElement.querySelectorAll('button');
    buttons.forEach(button => {
      button.replaceWith(button.cloneNode(true));
    });
  }

  /**
   * パフォーマンスメトリクスの取得
   */
  getPerformanceMetrics() {
    return {
      ...this.performanceMetrics,
      isInitialized: this.isInitialized,
      cachedModals: this.modals.size,
      activeListeners: this.eventListeners.size
    };
  }

  /**
   * 年齢層の設定
   * @param {string} ageGroup - 年齢層
   */
  setAgeGroup(ageGroup) {
    if (Object.values(this.config.AGE_GROUPS).includes(ageGroup)) {
      this.currentAge = ageGroup;
      console.log(`[ModalManager] Age group set to: ${ageGroup}`);
    } else {
      console.warn(`[ModalManager] Invalid age group: ${ageGroup}`);
    }
  }

  /**
   * リソースのクリーンアップ
   */
  cleanup() {
    // 全てのイベントリスナーをクリーンアップ
    this.eventListeners.forEach((listener, element) => {
      element.removeEventListener(listener.event, listener.handler);
    });
    this.eventListeners.clear();
    
    // モーダル要素をクリーンアップ
    this.modals.forEach(modal => {
      if (modal.parentNode) {
        modal.parentNode.removeChild(modal);
      }
    });
    this.modals.clear();
    
    this.isInitialized = false;
    console.log('[ModalManager] Cleanup completed');
  }
}

// シングルトンインスタンス
const unifiedModalManager = new UnifiedModalManager();

/**
 * グローバル関数（後方互換性）
 */
function showPrivacyModal(onContinue, options = {}) {
  return unifiedModalManager.showPrivacyModal({
    onContinue: onContinue,
    ...options
  });
}

function showDigitalCitizenshipModal(onAccept, options = {}) {
  return unifiedModalManager.showDigitalCitizenshipModal({
    onAccept: onAccept,
    ...options
  });
}

/**
 * 初期化関数
 */
function initializeModalSystem(options = {}) {
  return unifiedModalManager.initialize(options);
}

/**
 * 年齢層設定関数
 */
function setModalAgeGroup(ageGroup) {
  return unifiedModalManager.setAgeGroup(ageGroup);
}

// UnifiedModalManagerクラスに追加の補助メソッドを追加
UnifiedModalManager.prototype.updatePerformanceMetrics = function(action, value) {
  try {
    this.performanceMetrics[action] = (this.performanceMetrics[action] || 0) + 1;
    if (value && action.includes('time')) {
      this.performanceMetrics.avgLoadTime = (this.performanceMetrics.avgLoadTime + value) / 2;
    }
  } catch (error) {
    console.warn('[ModalManager] Performance metrics update failed:', error);
  }
};

UnifiedModalManager.prototype.createModalPromiseWithTimeout = function(modalElement, options, operationId) {
  return new Promise((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      console.error(`[ModalManager:${operationId}] Modal timeout`);
      reject(new Error('モーダル操作がタイムアウトしました'));
    }, 30000); // 30秒タイムアウト

    this.setupModalResultHandlers(modalElement, options, (result) => {
      clearTimeout(timeoutId);
      resolve(result);
    }, (error) => {
      clearTimeout(timeoutId);
      reject(error);
    });
  });
};

UnifiedModalManager.prototype.handleModalError = function(error, operationId, options) {
  this.performanceMetrics.errorCount++;
  console.error(`[ModalManager:${operationId}] Error:`, error);
  
  // エラー情報をローカルストレージに記録（デバッグ用）
  try {
    const errorLog = {
      timestamp: new Date().toISOString(),
      operationId: operationId,
      error: error.message,
      options: options
    };
    const existingLogs = JSON.parse(localStorage.getItem('modalManager_errorLogs') || '[]');
    existingLogs.push(errorLog);
    // 最新50件のみ保持
    if (existingLogs.length > 50) {
      existingLogs.splice(0, existingLogs.length - 50);
    }
    localStorage.setItem('modalManager_errorLogs', JSON.stringify(existingLogs));
  } catch (logError) {
    console.warn('[ModalManager] Error logging failed:', logError);
  }
};

UnifiedModalManager.prototype.createUserFriendlyError = function(originalError, context) {
  const userFriendlyMessages = {
    'モーダル表示': 'モーダルウィンドウの表示中にエラーが発生しました。ページを再読み込みしてもう一度お試しください。',
    'コンテンツ取得': 'コンテンツの読み込み中にエラーが発生しました。インターネット接続を確認してください。',
    '要素作成': 'ページの構築中にエラーが発生しました。ブラウザを更新してください。'
  };
  
  const friendlyMessage = userFriendlyMessages[context] || 'システムエラーが発生しました。';
  const error = new Error(friendlyMessage);
  error.originalError = originalError;
  return error;
};

UnifiedModalManager.prototype.getFallbackContentSafe = function(type) {
  try {
    const fallbackContent = getFallbackContent(type);
    return fallbackContent;
  } catch (error) {
    console.error('[ModalManager] Fallback content creation failed:', error);
    return {
      title: 'システム情報',
      content: '<p>コンテンツを読み込んでいます...</p>',
      actionText: 'OK'
    };
  }
};

UnifiedModalManager.prototype.buildModalElement = function(type, content, options) {
  const modal = document.createElement('div');
  modal.className = 'fixed inset-0 bg-black/80 z-[60] flex items-center justify-center p-4';
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');
  
  const modalContent = document.createElement('div');
  modalContent.className = 'glass-panel rounded-xl p-6 max-w-4xl w-full max-h-[90vh] overflow-y-auto';
  
  modalContent.innerHTML = `
    <div class="flex items-center justify-between mb-6">
      <h3 class="text-2xl font-bold text-cyan-400">${this.escapeHtml(content.title)}</h3>
      <button type="button" class="modal-close p-2 text-gray-400 transition-colors rounded-full hover:bg-gray-700 hover:text-white focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-cyan-500" aria-label="閉じる">
        <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
        </svg>
      </button>
    </div>
    
    <div class="modal-body text-gray-300">
      ${content.content}
    </div>
    
    <div class="flex justify-end gap-3 mt-6">
      <button type="button" class="modal-cancel btn btn-secondary">キャンセル</button>
      <button type="button" class="modal-action btn btn-primary">
        ${this.escapeHtml(content.actionText || '続行')}
      </button>
    </div>
  `;
  
  modal.appendChild(modalContent);
  return modal;
};

UnifiedModalManager.prototype.setupAccessibility = function(modalElement) {
  try {
    // フォーカストラップの設定
    const focusableElements = modalElement.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
    const firstFocusable = focusableElements[0];
    const lastFocusable = focusableElements[focusableElements.length - 1];
    
    if (firstFocusable) {
      firstFocusable.focus();
    }
    
    // Tab キーでフォーカスがループするように
    modalElement.addEventListener('keydown', (e) => {
      if (e.key === 'Tab') {
        if (e.shiftKey) {
          if (document.activeElement === firstFocusable) {
            e.preventDefault();
            lastFocusable.focus();
          }
        } else {
          if (document.activeElement === lastFocusable) {
            e.preventDefault();
            firstFocusable.focus();
          }
        }
      }
    });
  } catch (error) {
    console.warn('[ModalManager] Accessibility setup failed:', error);
  }
};

UnifiedModalManager.prototype.setupModalResultHandlers = function(modalElement, options, resolve, reject) {
  try {
    const handleClose = () => {
      this.closeModalSafely(modalElement);
      resolve({ action: 'cancelled' });
    };
    
    const handleAction = () => {
      this.closeModalSafely(modalElement);
      resolve({ action: 'confirmed' });
      
      // コールバック実行
      if (options.onContinue) options.onContinue();
      if (options.onAccept) options.onAccept();
    };
    
    // イベントリスナーの設定
    modalElement.querySelector('.modal-close')?.addEventListener('click', handleClose);
    modalElement.querySelector('.modal-cancel')?.addEventListener('click', handleClose);
    modalElement.querySelector('.modal-action')?.addEventListener('click', handleAction);
    
    // ESCキーでの終了
    const handleEsc = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        handleClose();
        document.removeEventListener('keydown', handleEsc);
      }
    };
    document.addEventListener('keydown', handleEsc);
    
    // オーバーレイクリックでの終了
    modalElement.addEventListener('click', (e) => {
      if (e.target === modalElement) {
        handleClose();
      }
    });
    
  } catch (error) {
    console.error('[ModalManager] Result handler setup failed:', error);
    reject(error);
  }
};

UnifiedModalManager.prototype.closeModalSafely = function(modalElement) {
  try {
    // フェードアウトアニメーション
    modalElement.style.opacity = '0';
    
    setTimeout(() => {
      if (modalElement.parentNode) {
        this.cleanupModalEventListeners(modalElement);
        modalElement.remove();
      }
    }, this.config.ANIMATION.FADE_DURATION);
    
  } catch (error) {
    console.error('[ModalManager] Modal close failed:', error);
    // フォールバック: 即座に削除
    if (modalElement.parentNode) {
      modalElement.remove();
    }
  }
};

/**
 * フロントエンドからのモーダルコンテンツ取得用関数
 * @param {string} modalType モーダルの種類
 * @param {string} ageGroup 年齢グループ
 * @return {Object} モーダルコンテンツオブジェクト
 */
function getModalContent(modalType, ageGroup) {
  try {
    // キャッシュからの取得を試行
    const cacheKey = `modal_${modalType}_${ageGroup}`;
    const cache = CacheService.getScriptCache();
    const cachedContent = cache.get(cacheKey);
    
    if (cachedContent) {
      console.log(`✅ モーダルコンテンツをキャッシュから取得: ${modalType}/${ageGroup}`);
      return JSON.parse(cachedContent);
    }
    
    // テンプレートから取得
    const content = MODAL_CONTENT_TEMPLATES[modalType]?.[ageGroup];
    if (!content) {
      console.warn(`⚠️ モーダルコンテンツが見つかりません: ${modalType}/${ageGroup}`);
      // フォールバックコンテンツを返す
      return getFallbackContent(modalType);
    }
    
    // キャッシュに保存（5分間）
    cache.put(cacheKey, JSON.stringify(content), MODAL_CONFIG.CACHE_TTL);
    
    console.log(`✅ モーダルコンテンツを生成: ${modalType}/${ageGroup}`);
    return content;
    
  } catch (error) {
    console.error('❌ モーダルコンテンツ取得エラー:', error);
    return getFallbackContent(modalType);
  }
}

/**
 * フォールバックコンテンツの取得
 * @param {string} modalType モーダルの種類
 * @return {Object} フォールバックコンテンツ
 */
function getFallbackContent(modalType) {
  const fallbacks = {
    'privacy-policy': {
      title: 'プライバシーとデータ保護',
      content: `
        <div class="space-y-4">
          <p>このシステムは教育目的でのみ使用されます。</p>
          <ul class="list-disc list-inside space-y-1">
            <li>個人情報は適切に保護されます</li>
            <li>回答データは教育目的でのみ使用されます</li>
            <li>管理者は適切なアクセス権限を持っています</li>
            <li>データは法的要件に従って管理されます</li>
          </ul>
        </div>
      `,
      actionText: '理解しました、継続'
    },
    'digital-citizenship': {
      title: 'デジタル・シティズンシップ',
      content: `
        <div class="space-y-4">
          <p>デジタル社会における責任ある市民として、技術を適切かつ効果的に使用しましょう。</p>
          <ul class="list-disc list-inside space-y-1">
            <li>他者への配慮と尊重を心がける</li>
            <li>建設的で前向きなコミュニケーション</li>
            <li>情報を批判的に評価する</li>
            <li>デジタル環境でのマナーとルールを守る</li>
          </ul>
        </div>
      `,
      actionText: '理解しました'
    }
  };
  
  return fallbacks[modalType] || {
    title: 'システム情報',
    content: '<p>コンテンツを読み込み中です...</p>',
    actionText: 'OK'
  };
}

/**
 * モーダル表示のログ記録（分析用）
 * @param {string} userId ユーザーID
 * @param {string} modalType モーダルの種類
 * @param {string} ageGroup 年齢グループ
 * @param {string} action 実行されたアクション
 */
function logModalInteraction(userId, modalType, ageGroup, action) {
  try {
    const logData = {
      timestamp: new Date().toISOString(),
      userId: userId,
      modalType: modalType,
      ageGroup: ageGroup,
      action: action, // 'shown', 'accepted', 'cancelled'
      userAgent: Session.getTemporaryActiveUserKey() || 'unknown'
    };
    
    // ログをSpreadsheetに記録（オプション機能）
    const logSheet = getModalLogSheet();
    if (logSheet) {
      logSheet.appendRow([
        logData.timestamp,
        logData.userId,
        `modal_${logData.modalType}`,
        `${logData.ageGroup}_${logData.action}`,
        logData.userAgent
      ]);
    }
    
    console.log(`📊 モーダル操作ログ: ${modalType}/${ageGroup}/${action}`);
    
  } catch (error) {
    console.error('❌ モーダルログ記録エラー:', error);
  }
}

/**
 * ログ記録用シートの取得（存在しない場合は作成しない）
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getModalLogSheet() {
  try {
    const databaseId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (!databaseId) return null;
    
    const database = SpreadsheetApp.openById(databaseId);
    return database.getSheetByName('ModalLogs');
  } catch (error) {
    return null;
  }
}