/**
 * プロダクション形式テストデータフィクスチャ
 * 実際の運用環境で収集されたデータ形式を再現
 */

import { MockTestUtils } from '../mocks/gasMocks';

// Real production data structures
export interface ProductionUserData {
  userId: string;
  adminEmail: string;
  spreadsheetId: string;
  spreadsheetUrl: string;
  createdAt: string;
  configJson: string;
  lastAccessedAt: string;
  isActive: string;
}

export interface ProductionFormResponse {
  timestamp: string;
  email: string;
  answer: string;
  reason?: string;
  class?: string;
  name?: string;
  reactions: {
    understand: number;
    like: number;
    curious: number;
  };
}

export interface ProductionSpreadsheetData {
  headers: string[];
  rows: any[][];
  metadata: {
    totalResponses: number;
    dateRange: {
      start: string;
      end: string;
    };
    schools: string[];
    classes: string[];
  };
}

// Realistic production data sets
export const PRODUCTION_DATA_FIXTURES = {
  // Typical Japanese school form responses (anonymized)
  SCHOOL_SURVEY_2024: {
    headers: [
      'タイムスタンプ',
      'メールアドレス',
      '将来の夢について教えてください',
      'なぜそう思うのか理由を教えてください',
      'あなたのクラス',
      'お名前（任意）',
      'なるほど！',
      'いいね！',
      'もっと知りたい！',
    ],
    responses: [
      [
        '2024/01/15 9:23:45',
        'student001@tokyo-edu.ac.jp',
        'AIエンジニアになって、人々の生活をより便利にしたいです。',
        '最近のAI技術の進歩を見ていて、無限の可能性を感じるからです。',
        '3年A組',
        '田中',
        '12',
        '8',
        '15',
      ],
      [
        '2024/01/15 9:45:12',
        'student002@tokyo-edu.ac.jp',
        '環境問題の解決に取り組む研究者になりたいです。',
        '地球温暖化などの問題が深刻化しており、未来のために何かしたいからです。',
        '3年B組',
        '佐藤',
        '25',
        '18',
        '22',
      ],
      [
        '2024/01/15 10:12:33',
        'student003@tokyo-edu.ac.jp',
        '医師になって、特に小児科で働きたいと思っています。',
        '子供の時に病気で入院した際、優しい先生に支えられた経験があるからです。',
        '3年A組',
        '鈴木',
        '30',
        '25',
        '18',
      ],
      [
        '2024/01/15 10:34:21',
        'student004@tokyo-edu.ac.jp',
        '国際的に活躍できる外交官を目指しています。',
        '世界平和に貢献し、異なる文化を繋ぐ架け橋になりたいからです。',
        '2年C組',
        '山田',
        '15',
        '20',
        '28',
      ],
      [
        '2024/01/15 11:01:45',
        'student005@tokyo-edu.ac.jp',
        'ゲームクリエイターになって、世界中の人を楽しませたいです。',
        'ゲームを通じて人と人を繋げ、新しい体験を提供できると思うからです。',
        '2年A組',
        '高橋',
        '45',
        '38',
        '42',
      ],
      [
        '2024/01/15 11:23:18',
        'student006@tokyo-edu.ac.jp',
        '建築家になって、人々が快適に暮らせる住空間を作りたいです。',
        '美しく機能的な建物が人の心を豊かにする力があると信じているからです。',
        '1年B組',
        '伊藤',
        '22',
        '16',
        '25',
      ],
    ],
  },

  // Corporate training feedback data
  CORPORATE_TRAINING_2024: {
    headers: [
      'タイムスタンプ',
      'メールアドレス',
      '今日の研修で最も印象に残ったポイントは何ですか？',
      'そう感じた理由を具体的に教えてください',
      '所属部署',
      'お名前',
      '参考になった',
      '実践したい',
      '他にも知りたい',
    ],
    responses: [
      [
        '2024/03/15 16:45:30',
        'yamada@company.co.jp',
        'チームコミュニケーションの重要性について深く理解できました。',
        '実際のケーススタディを通じて、効果的な対話の方法を学べたからです。',
        '営業部',
        '山田太郎',
        '8',
        '9',
        '6',
      ],
      [
        '2024/03/15 16:52:15',
        'tanaka@company.co.jp',
        'プロジェクト管理ツールの活用法が非常に参考になりました。',
        '現在進行中のプロジェクトで即座に応用できそうな内容だったからです。',
        'IT部',
        '田中花子',
        '10',
        '10',
        '8',
      ],
    ],
  },

  // Academic conference feedback
  ACADEMIC_CONFERENCE_2024: {
    headers: [
      'Timestamp',
      'Email Address',
      'Most valuable session',
      'Why was it valuable?',
      'Institution',
      'Name (Optional)',
      'Informative',
      'Applicable',
      'Want to learn more',
    ],
    responses: [
      [
        '2024/05/20 14:30:22',
        'prof.smith@university.edu',
        'AI Ethics in Research - A comprehensive overview of ethical considerations',
        'The session provided practical frameworks for implementing ethical guidelines in AI research projects.',
        'Tokyo University',
        'Dr. Smith',
        '15',
        '12',
        '18',
      ],
      [
        '2024/05/20 14:45:33',
        'researcher@institute.ac.jp',
        'Machine Learning in Healthcare - Real-world applications and challenges',
        'The case studies demonstrated clear pathways from research to clinical implementation.',
        '医療AI研究所',
        '研究員A',
        '20',
        '25',
        '22',
      ],
    ],
  },

  // Edge case problematic data
  PROBLEMATIC_DATA_2024: {
    headers: [
      'タイムスタンプ',
      'メールアドレス',
      '意見',
      '理由',
      'クラス',
      '名前',
      'なるほど',
      'いいね',
      'もっと知りたい',
    ],
    responses: [
      // Empty/minimal data
      ['2024/01/15 12:00:00', 'empty@test.com', '', '', '', '', '0', '0', '0'],
      // Very long text
      [
        '2024/01/15 12:05:00',
        'longtext@test.com',
        'これは非常に長い意見です。'.repeat(100),
        'これは非常に長い理由です。'.repeat(50),
        'クラス名が長い場合のテスト用クラス名'.repeat(10),
        '名前が非常に長い場合のテスト用名前'.repeat(5),
        '999',
        '999',
        '999',
      ],
      // Special characters and injection attempts
      [
        '2024/01/15 12:10:00',
        'special@test.com',
        '<script>alert("XSS")</script>特殊文字テスト🎌',
        "'; DROP TABLE users; -- SQL injection attempt",
        '3年<script>alert("XSS")</script>組',
        'テスト\n改行\t\tタブ\\バックスラッシュ',
        'NaN',
        'Infinity',
        '-1',
      ],
      // Unicode and internationalization
      [
        '2024/01/15 12:15:00',
        'unicode@test.com',
        'مرحبا بكم في العالم العربي 🌍',
        '这是中文测试文本 中國語テスト',
        '한국어 클래스',
        'Üñïçødé Ñâmé',
        '∞',
        '∑',
        '∫',
      ],
    ],
  },
};

// Production-like database configurations
export const PRODUCTION_DATABASE_FIXTURES = {
  USERS: [
    {
      userId: 'usr_2024_001_tokyo_hs',
      adminEmail: 'teacher001@tokyo-edu.ac.jp',
      spreadsheetId: '1ABC123DEF456GHI789JKL012MNO345PQR678STU',
      spreadsheetUrl:
        'https://docs.google.com/spreadsheets/d/1ABC123DEF456GHI789JKL012MNO345PQR678STU/edit',
      createdAt: '2024-01-10T09:00:00.000Z',
      configJson: JSON.stringify({
        appPublished: true,
        publishedSheetName: 'フォームの回答 1',
        setupStatus: 'completed',
        formCreated: true,
        formUrl: 'https://forms.google.com/forms/d/1234567890abcdef/viewform',
        lastPublished: '2024-01-15T10:30:00.000Z',
        surveyTitle: '将来の夢についてのアンケート',
        maxResponses: 1000,
        autoStop: false,
        displaySettings: {
          showNames: true,
          showReactions: true,
          anonymousMode: false,
        },
      }),
      lastAccessedAt: '2024-01-15T11:45:22.000Z',
      isActive: 'true',
    },
    {
      userId: 'usr_2024_002_corp_train',
      adminEmail: 'hr.admin@company.co.jp',
      spreadsheetId: '1XYZ789ABC012DEF345GHI678JKL901MNO234PQR',
      spreadsheetUrl:
        'https://docs.google.com/spreadsheets/d/1XYZ789ABC012DEF345GHI678JKL901MNO234PQR/edit',
      createdAt: '2024-03-01T08:30:00.000Z',
      configJson: JSON.stringify({
        appPublished: true,
        publishedSheetName: '研修フィードバック',
        setupStatus: 'completed',
        formCreated: true,
        formUrl: 'https://forms.google.com/forms/d/corporate_training/viewform',
        lastPublished: '2024-03-15T16:00:00.000Z',
        surveyTitle: '社内研修フィードバック',
        maxResponses: 500,
        autoStop: true,
        autoStopDate: '2024-03-31T23:59:59.000Z',
        displaySettings: {
          showNames: false,
          showReactions: true,
          anonymousMode: true,
        },
      }),
      lastAccessedAt: '2024-03-15T17:30:15.000Z',
      isActive: 'true',
    },
    {
      userId: 'usr_2024_003_acad_conf',
      adminEmail: 'conference@university.edu',
      spreadsheetId: '1QWE456RTY789UIO123ASD456FGH789JKL012ZXC',
      spreadsheetUrl:
        'https://docs.google.com/spreadsheets/d/1QWE456RTY789UIO123ASD456FGH789JKL012ZXC/edit',
      createdAt: '2024-05-01T12:00:00.000Z',
      configJson: JSON.stringify({
        appPublished: true,
        publishedSheetName: 'Conference Feedback',
        setupStatus: 'completed',
        formCreated: true,
        formUrl: 'https://forms.google.com/forms/d/academic_conference/viewform',
        lastPublished: '2024-05-20T13:30:00.000Z',
        surveyTitle: 'Academic Conference Feedback 2024',
        maxResponses: 200,
        autoStop: false,
        displaySettings: {
          showNames: true,
          showReactions: true,
          anonymousMode: false,
        },
        language: 'en',
      }),
      lastAccessedAt: '2024-05-20T15:22:33.000Z',
      isActive: 'true',
    },
  ],
};

// Helper functions for test data setup
export class ProductionDataManager {
  /**
   * Setup realistic test environment with production-like data
   */
  static setupProductionEnvironment(dataSetName: keyof typeof PRODUCTION_DATA_FIXTURES): void {
    const dataset = PRODUCTION_DATA_FIXTURES[dataSetName];
    if (!dataset) {
      throw new Error(`Dataset ${dataSetName} not found`);
    }

    MockTestUtils.clearAllMockData();
    MockTestUtils.createRealisticSpreadsheetData(dataset.headers, dataset.responses);

    // Setup corresponding user data
    const userData =
      PRODUCTION_DATABASE_FIXTURES.USERS.find(
        (user) =>
          user.configJson.includes(dataSetName.toLowerCase()) ||
          (dataSetName.includes('SCHOOL') && user.adminEmail.includes('edu')) ||
          (dataSetName.includes('CORPORATE') && user.adminEmail.includes('company')) ||
          (dataSetName.includes('ACADEMIC') && user.adminEmail.includes('university'))
      ) || PRODUCTION_DATABASE_FIXTURES.USERS[0];

    MockTestUtils.setMockData('currentUser', userData);
    MockTestUtils.setMockData(`user_${userData.userId}`, userData);
  }

  /**
   * Generate large-scale test data for performance testing
   */
  static generateLargeDataset(
    responseCount: number,
    baseDataSet: keyof typeof PRODUCTION_DATA_FIXTURES
  ): ProductionSpreadsheetData {
    const dataset = PRODUCTION_DATA_FIXTURES[baseDataSet];
    if (!dataset) {
      throw new Error(`Dataset ${baseDataSet} not found`);
    }

    const baseResponse = dataset.responses[0];
    const generatedResponses: any[][] = [];
    const schools = ['Tokyo University', 'Osaka Institute', 'Kyoto Academy', 'Nagoya School'];
    const classes = ['1年A組', '1年B組', '2年A組', '2年B組', '3年A組', '3年B組'];

    for (let i = 0; i < responseCount; i++) {
      const timestamp = new Date(
        2024,
        0,
        15,
        9 + Math.floor(i / 60),
        i % 60,
        Math.floor(Math.random() * 60)
      );
      const response = [
        timestamp.toLocaleString('ja-JP'),
        `student${String(i).padStart(4, '0')}@test.edu`,
        `${baseResponse[2]} (Generated response ${i})`,
        `${baseResponse[3]} (Generated reason ${i})`,
        classes[i % classes.length],
        `生徒${i}`,
        String(Math.floor(Math.random() * 50)),
        String(Math.floor(Math.random() * 50)),
        String(Math.floor(Math.random() * 50)),
      ];
      generatedResponses.push(response);
    }

    return {
      headers: dataset.headers,
      rows: generatedResponses,
      metadata: {
        totalResponses: responseCount,
        dateRange: {
          start: '2024-01-15T09:00:00Z',
          end: new Date().toISOString(),
        },
        schools: schools,
        classes: classes,
      },
    };
  }

  /**
   * Setup edge case data for robustness testing
   */
  static setupEdgeCaseData(): void {
    MockTestUtils.clearAllMockData();
    MockTestUtils.createRealisticSpreadsheetData(
      PRODUCTION_DATA_FIXTURES.PROBLEMATIC_DATA_2024.headers,
      PRODUCTION_DATA_FIXTURES.PROBLEMATIC_DATA_2024.responses
    );
  }

  /**
   * Validate data integrity and consistency
   */
  static validateDataIntegrity(data: ProductionSpreadsheetData): boolean {
    // Header validation
    if (!data.headers || !Array.isArray(data.headers) || data.headers.length === 0) {
      return false;
    }

    // Row validation
    if (!data.rows || !Array.isArray(data.rows)) {
      return false;
    }

    // Consistency validation
    const headerCount = data.headers.length;
    const inconsistentRows = data.rows.filter(
      (row) => !Array.isArray(row) || row.length !== headerCount
    );

    if (inconsistentRows.length > 0) {
      console.warn(`Found ${inconsistentRows.length} inconsistent rows`);
      return false;
    }

    // Metadata validation
    if (data.metadata) {
      if (data.metadata.totalResponses !== data.rows.length) {
        console.warn('Metadata totalResponses does not match actual row count');
        return false;
      }
    }

    return true;
  }

  /**
   * Generate realistic column mapping based on headers
   */
  static generateColumnMapping(headers: string[]): Record<string, number> {
    const mapping: Record<string, number> = {
      timestamp: -1,
      email: -1,
      answer: -1,
      reason: -1,
      class: -1,
      name: -1,
      reactions: [],
    };

    headers.forEach((header, index) => {
      const lowerHeader = header.toLowerCase();

      if (lowerHeader.includes('タイムスタンプ') || lowerHeader.includes('timestamp')) {
        mapping.timestamp = index;
      } else if (lowerHeader.includes('メール') || lowerHeader.includes('email')) {
        mapping.email = index;
      } else if (
        lowerHeader.includes('意見') ||
        lowerHeader.includes('夢') ||
        lowerHeader.includes('session') ||
        lowerHeader.includes('印象')
      ) {
        mapping.answer = index;
      } else if (
        lowerHeader.includes('理由') ||
        lowerHeader.includes('why') ||
        lowerHeader.includes('感じた')
      ) {
        mapping.reason = index;
      } else if (
        lowerHeader.includes('クラス') ||
        lowerHeader.includes('部署') ||
        lowerHeader.includes('institution') ||
        lowerHeader.includes('class')
      ) {
        mapping.class = index;
      } else if (lowerHeader.includes('名前') || lowerHeader.includes('name')) {
        mapping.name = index;
      } else if (
        lowerHeader.includes('なるほど') ||
        lowerHeader.includes('いいね') ||
        lowerHeader.includes('知りたい') ||
        lowerHeader.includes('informative') ||
        lowerHeader.includes('applicable') ||
        lowerHeader.includes('learn')
      ) {
        (mapping.reactions as number[]).push(index);
      }
    });

    return mapping;
  }
}

// Export for use in tests
export default ProductionDataManager;
