'use strict';
// NEED: config.js
// NEED: GoogleAPIFunctions.js
// NEED: KintoneAPIFunctions.js
// NEED: ui.js

const currentConfig = window.config.production;

setButton('outputSaleSettingsStatement', 'セール価格設定リスト出力', executeSaleSettingsDataFetch, true, 'detail');

async function executeSaleSettingsDataFetch(updateStatus = () => {}) {
  try {
    const config = currentConfig.kintone.app;
    
    // ----------------------------------------------------
    // 0. 開いている「セール設定」レコードからパラメータを取得
    // ----------------------------------------------------
    const currentRecordObj = kintone.app.record.get();
    if (!currentRecordObj || !currentRecordObj.record) {
      throw new Error('レコード情報が取得できませんでした。詳細画面から実行してください。');
    }
    const currentRecord = currentRecordObj.record;
    
    const saleStartDate = currentRecord['セール開始日']?.value || '';
    const saleEndDate = currentRecord['セール終了日']?.value || '';
    
    // 親フォルダIDは config から取得 (ユーザーにて config.js に追加される想定)
    const targetFolderId = currentConfig.google.drive['セール設定']?.target_folder_id;
    if (!targetFolderId) {
      throw new Error('エラー: config.js の google.drive.セール設定.target_folder_id が設定されていません。');
    }

    // ----------------------------------------------------
    // 1. 各種データの一括事前ロード
    // ----------------------------------------------------
    // (1) 商品マスタ（販売中の商品）を一括取得
    updateStatus('商品マスタ（販売中の商品）取得中...');
    const allActiveProducts = await fetchRecordsFromApp(
      config['商品マスタ'].api_token,
      config['商品マスタ'].id,
      config['商品マスタ'].view['販売中の商品']
    );
    
    if (!allActiveProducts || allActiveProducts.length === 0) {
      throw new Error('販売中の商品が見つかりませんでした。');
    }

    // (2) 絞り込み用の各種キーの抽出（前後のスペースや改行を除去）
    const csCodes = [...new Set(allActiveProducts.map(r => r['CS別品番']?.value?.trim()?.replace(/\r?\n/g, '')).filter(Boolean))];
    const brandCodes = [...new Set(allActiveProducts.map(r => r['ブランド品番']?.value?.trim()?.replace(/\r?\n/g, '')).filter(Boolean))];
    const makerIds = [...new Set(allActiveProducts.map(r => (r['メーカーID']?.value || r['登録メーカーID']?.value)?.trim()?.replace(/\r?\n/g, '')).filter(Boolean))];

    // (3) 各種データの一括ロード
    
    // (A) 在庫データの分割クエリによる取得 (対象品番のみに制限)
    updateStatus('在庫データ取得中...');
    const clientInv = new KintoneRestAPIClient({ auth: { apiToken: config['在庫'].api_token } });
    const allInventoryData = await fetchRecordsInChunks(clientInv, config['在庫'].id, csCodes, 'CS品番');

    // (B) 在庫分析データの分割クエリによる取得 (対象品番のみに制限)
    updateStatus('在庫分析データ取得中...');
    const clientAnalysis = new KintoneRestAPIClient({ auth: { apiToken: config['在庫分析'].api_token } });
    const allInventoryAnalysisData = await fetchRecordsInChunks(clientAnalysis, config['在庫分析'].id, csCodes, 'CS品番');

    // (C) salegoodsデータの分割クエリによる取得 (対象ブランドのみに制限)
    updateStatus('salegoodsデータ取得中...');
    const clientSale = new KintoneRestAPIClient({ auth: { apiToken: config['salegoods'].api_token } });
    const allSaleGoodsData = await fetchRecordsInChunks(clientSale, config['salegoods'].id, brandCodes, 'ブランド品番');

    // (D) メーカーマスタ & 親カテゴリマスタの一括ロード（ビュー条件の動的 and 結合）
    updateStatus('メーカーマスタおよび親カテゴリマスタ取得中...');
    let makerQuery = makerIds.length > 0 ? `メーカーID in (${makerIds.map(v => '"' + String(v).replace(/"/g, '\\"') + '"').join(',')})` : '';
    if (makerQuery) {
      const v = await (new KintoneRestAPIClient({ auth: { apiToken: config['メーカーマスタ from SQL'].api_token } })).app.getViews({ app: config['メーカーマスタ from SQL'].id });
      if (Object.values(v.views).find(w => w.id === config['メーカーマスタ from SQL'].view['連絡先'])?.filterCond) makerQuery = ' and ' + makerQuery;
    }
    const allManufacturers = await fetchRecordsFromApp(
      config['メーカーマスタ from SQL'].api_token,
      config['メーカーマスタ from SQL'].id,
      config['メーカーマスタ from SQL'].view['連絡先'],
      makerQuery
    );
    
    let parentQuery = makerIds.length > 0 ? `メーカーコード in (${makerIds.map(v => '"' + String(v).replace(/"/g, '\\"') + '"').join(',')})` : '';
    if (parentQuery) {
      const v = await (new KintoneRestAPIClient({ auth: { apiToken: config['親カテゴリマスタ'].api_token } })).app.getViews({ app: config['親カテゴリマスタ'].id });
      if (Object.values(v.views).find(w => w.id === config['親カテゴリマスタ'].view['連絡先'])?.filterCond) parentQuery = ' and ' + parentQuery;
    }
    const allParentCategories = await fetchRecordsFromApp(
      config['親カテゴリマスタ'].api_token,
      config['親カテゴリマスタ'].id,
      config['親カテゴリマスタ'].view['連絡先'],
      parentQuery
    );

    // (D-2) 子カテゴリマスタの一括ロード
    updateStatus('子カテゴリマスタ取得中...');
    const clientSubCategory = new KintoneRestAPIClient({ auth: { apiToken: config['子カテゴリマスタ'].api_token } });
    const allSubCategories = await clientSubCategory.record.getAllRecords({ app: config['子カテゴリマスタ'].id });

    // (D-3) ZOZO商品タイプマスタの一括ロード
    updateStatus('ZOZO商品タイプマスタ取得中...');
    const clientZozoType = new KintoneRestAPIClient({ auth: { apiToken: config['ZOZO商品タイプマスタ'].api_token } });
    const allZozoTypes = await clientZozoType.record.getAllRecords({ app: config['ZOZO商品タイプマスタ'].id });

    // (E) Z.売上・注文 データの取得（日付条件なしで全件取得）
    updateStatus('売上・注文データ取得中...');
    const clientSales = new KintoneRestAPIClient({ auth: { apiToken: config['Z_売上（注文日）'].api_token } });
    const allSalesData = await clientSales.record.getAllRecords({
      app: config['Z_売上（注文日）'].id
    });
    
    // (8) Googleアクセストークンの取得
    updateStatus('Googleアクセストークン取得中...');
    let accessToken = await getAccessToken();
    
    // (9) Google Drive上での日付フォルダ（yyyy/mm/dd）の作成 (指定された親フォルダ配下)
    const templateSpreadsheetId = currentConfig.google.drive['セール設定']?.template_spreadsheet_id;
    if (!templateSpreadsheetId) {
      throw new Error('エラー: config.js の google.drive.セール設定.template_spreadsheet_id が設定されていません。');
    }
    
    const today = new Date();
    const formattedDateSlash = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
    const formattedDateNoSlash = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    
    updateStatus(`日付フォルダ（${formattedDateSlash}）作成中...`);
    const dateFolder = await createDriveFolder(accessToken, targetFolderId, formattedDateSlash);
    const dateFolderId = dateFolder.id;
    const dateFolderUrl = dateFolder.webViewLink;

    // 現在開いているセール設定レコードの「セール設定価格リストフォルダ」フィールドを作成したフォルダのURLで更新
    updateStatus('レコードのフォルダURLを更新中...');
    const recordId = kintone.app.record.getId();
    const updateRecordBody = {
      app: kintone.app.getId(),
      id: recordId,
      record: {
        "セール設定価格リストフォルダ": { value: dateFolderUrl || dateFolderId }
      }
    };
    await kintone.api(kintone.api.url('/k/v1/record.json', true), 'PUT', updateRecordBody);
    
    // ----------------------------------------------------
    // 1.5 各種データの高速検索用インデックス（Map）を作成
    // ----------------------------------------------------
    updateStatus('インデックス作成中...');
    
    const inventoryMapByCs = {};
    allInventoryData.forEach(inv => {
      const cs = inv['CS品番']?.value?.trim()?.replace(/\r?\n/g, '');
      if (cs) {
        if (!inventoryMapByCs[cs]) inventoryMapByCs[cs] = [];
        inventoryMapByCs[cs].push(inv);
      }
    });

    const inventoryAnalysisMapByCs = {};
    allInventoryAnalysisData.forEach(an => {
      const cs = an['CS品番']?.value?.trim()?.replace(/\r?\n/g, '');
      if (cs) {
        if (!inventoryAnalysisMapByCs[cs]) inventoryAnalysisMapByCs[cs] = [];
        inventoryAnalysisMapByCs[cs].push(an);
      }
    });

    const saleGoodsMapByBrand = {};
    allSaleGoodsData.forEach(sg => {
      const brand = sg['ブランド品番']?.value?.trim()?.replace(/\r?\n/g, '');
      if (brand) {
        if (!saleGoodsMapByBrand[brand]) saleGoodsMapByBrand[brand] = [];
        saleGoodsMapByBrand[brand].push(sg);
      }
    });

    const salesMapByCs = {};
    allSalesData.forEach(sales => {
      const cs = sales['CS品番']?.value?.trim()?.replace(/\r?\n/g, '');
      if (cs) {
        if (!salesMapByCs[cs]) salesMapByCs[cs] = [];
        salesMapByCs[cs].push(sales);
      }
    });

    const subCategoryMapById = {};
    allSubCategories.forEach(sub => {
      const subId = sub['子カテゴリID']?.value?.trim();
      if (subId) {
        subCategoryMapById[subId] = sub;
      }
    });

    // 親カテゴリマスタのインデックス（キー: 親子複合コード）
    const parentCategoryMapByCompositeKey = {};
    allParentCategories.forEach(pc => {
      const compKey = pc['親子複合コード']?.value?.trim();
      if (compKey) {
        parentCategoryMapByCompositeKey[compKey] = pc;
      }
    });
    
    // ZOZO商品タイプマッピング（キー: 商品タイプID）
    const zozoTypeMapById = {};
    allZozoTypes.forEach(record => {
      const typeId = record['商品タイプID']?.value;
      if (typeId) {
        zozoTypeMapById[String(typeId).trim()] = {
          parentType: record['商品タイプカテゴリ名']?.value || '',
          childType: record['商品タイプ名']?.value || ''
        };
      }
    });
    
    // ----------------------------------------------------
    // 2. 取得した「販売中の商品」をグループ化（出力単位に応じてメーカー単位 or 親カテゴリ単位）
    // ----------------------------------------------------
    const productsByGroup = {};
    const groupMeta = {}; // 各グループのメタ情報
    
    allActiveProducts.forEach(product => {
      const makerCode = (product['メーカーID']?.value || product['登録メーカーID']?.value || '')?.trim()?.replace(/\r?\n/g, '');
      if (!makerCode) return; // 不明なものはスキップ
      
      // 商品マスタの子カテゴリID（または子カテゴリコード）を取得し、子カテゴリマスタから親子複合コードを特定
      const childCategoryId = (product['子カテゴリID']?.value || product['子カテゴリコード']?.value || product['子カテゴリ']?.value || '')?.trim();
      const subCategory = subCategoryMapById[childCategoryId];
      const compositeKey = subCategory?.['親子複合コード']?.value?.trim() || '';

      // 親子複合コードをキーに親カテゴリマスタから正確なレコードを取得
      const parentCategoryRecord = (compositeKey && (parentCategoryMapByCompositeKey[compositeKey] || allParentCategories.find(pc => pc['親子複合コード']?.value == compositeKey)))
        || allParentCategories.find(pc => pc['メーカーコード']?.value == makerCode);

      const mId = parentCategoryRecord?.['メーカーID']?.value || makerCode;
      const makerRecord = allManufacturers.find(m => m['メーカーID']?.value == mId);
      
      // 出力単位の判定（デフォルトは親カテゴリ）
      const outputUnit = makerRecord?.['セールリスト出力単位']?.value || '親カテゴリ';
      
      // 出力単位が「親カテゴリ」の場合は「親子複合コード」、出力単位が「メーカー」の場合は「メーカーID」とする
      const groupKey = outputUnit === 'メーカー' ? mId : (compositeKey || `${makerCode}_unknown`);
      
      if (!productsByGroup[groupKey]) {
        productsByGroup[groupKey] = [];
        groupMeta[groupKey] = {
          outputUnit,
          makerId: mId,
          makerCode,
          compositeKey,
          makerRecord,
          parentCategoryRecord
        };
      }
      productsByGroup[groupKey].push(product);
    });
    
    // ----------------------------------------------------
    // 3. グループごとのループ処理（メモリ上での超高速フィルタリング）
    // ----------------------------------------------------
    const groupKeys = Object.keys(productsByGroup);
    const totalGroups = groupKeys.length;
    
    let successCount = 0;
    let failCount = 0;
    let serialNumber = 0; // 今回の実行バッチ内で登録するレコード用の通し番号（連番）
    const failedManufacturers = [];
    const recordsToRegister = []; // kintone一括登録用バッファ
    
    for (let i = 0; i < totalGroups; i++) {
      const groupKey = groupKeys[i];
      const activeProducts = productsByGroup[groupKey];
      const meta = groupMeta[groupKey];
      
      const makerRecord = meta.makerRecord;
      const parentCategoryRecord = meta.parentCategoryRecord;
      
      const makerName = meta.outputUnit === 'メーカー'
        ? (makerRecord?.['メーカー名']?.value || activeProducts[0]['メーカー名']?.value || 'メーカー不明')
        : (parentCategoryRecord?.['メーカー名']?.value || activeProducts[0]['メーカー名']?.value || 'メーカー不明');
      
      const brandName = parentCategoryRecord?.['メーカー名']?.value || activeProducts[0]['メーカー名']?.value || 'ブランド名不明';
      // 親カテゴリマスタから親カテゴリコード（表記ゆれ対応）を取得し、メーカー単位の場合は一律 "1"、親カテゴリ単位の場合はマスタのコードをセットする
      const pCodeFromMaster = parentCategoryRecord?.['親カテゴリコード']?.value || parentCategoryRecord?.['親カテゴリーコード']?.value || '';
      const finalParentCategoryCode = meta.outputUnit === 'メーカー' ? '1' : (pCodeFromMaster || '1');

      const progressLabel = `(${i + 1}/${totalGroups}) ${makerName}`;
      
      const maxRetries = 4; // 最初の試行(1) + 自動リトライ(3)で計4回試行
      let attempt = 0;
      let isSuccess = false;
      let copiedSpreadsheet = null; // リトライ時のファイル重複作成を防止するためループ外で管理
      
      while (attempt < maxRetries && !isSuccess) {
        attempt++;
        try {
          // Googleアクセストークンを処理直前に毎回更新し、1時間の有効期限切れを防止
          accessToken = await getAccessToken();
          
          const currentCsCodes = [...new Set(activeProducts.map(r => r['CS別品番']?.value?.trim()?.replace(/\r?\n/g, '')).filter(Boolean))];
          const currentBrandCodes = [...new Set(activeProducts.map(r => r['ブランド品番']?.value?.trim()?.replace(/\r?\n/g, '')).filter(Boolean))];
          
          if (currentCsCodes.length === 0) {
            isSuccess = true;
            continue;
          }
          
          // ① 各種詳細データをメモリ上のインデックスから高速抽出
          const inventoryData = [];
          currentCsCodes.forEach(cs => {
            const records = inventoryMapByCs[cs];
            if (records) inventoryData.push(...records);
          });

          const inventoryAnalysisData = [];
          currentCsCodes.forEach(cs => {
            const records = inventoryAnalysisMapByCs[cs];
            if (records) inventoryAnalysisData.push(...records);
          });

          const saleGoodsData = [];
          currentBrandCodes.forEach(brand => {
            const records = saleGoodsMapByBrand[brand];
            if (records) saleGoodsData.push(...records);
          });
          
          // Z_売上（注文日）のこのグループの品番に該当する売上データを抽出
          const salesData = [];
          currentCsCodes.forEach(cs => {
            const records = salesMapByCs[cs];
            if (records) salesData.push(...records);
          });
          
          // セール期間内の受注数および売上額の集計 (CS品番ごと、価格タイプが「セール」のものに限定)
          const salesSummaryByCs = {};
          const salesAmountSummaryByCs = {};
          salesData.forEach(record => {
            const csCode = record['CS品番']?.value;
            const priceType = record['価格タイプ']?.value;
            
            // 価格タイプが「セール」のものだけを集計
            if (priceType !== 'セール') {
              return;
            }
            
            const quantity = Number(record['注文数']?.value || 0);
            const amount = Number(record['合計金額_税抜_']?.value || 0);
            
            if (csCode) {
              salesSummaryByCs[csCode] = (salesSummaryByCs[csCode] || 0) + quantity;
              salesAmountSummaryByCs[csCode] = (salesAmountSummaryByCs[csCode] || 0) + amount;
            }
          });
          
          updateStatus('処理中... ' + progressLabel);
          console.log(`--- [セール設定処理開始] ${progressLabel} ---`);
          const productsByBrand = {};
          activeProducts.forEach(product => {
            const brandCode = product['ブランド品番']?.value;
            if (!brandCode) return; // ブランド品番がないものはスキップ
            if (!productsByBrand[brandCode]) {
              productsByBrand[brandCode] = [];
            }
            productsByBrand[brandCode].push(product);
          });
          
          const excelRows = Object.keys(productsByBrand).map(brandCode => {
            const brandProducts = productsByBrand[brandCode];
            const representativeProduct = brandProducts[0];
            const csCodesOfBrand = brandProducts.map(p => p['CS別品番']?.value).filter(Boolean);
            
            // 在庫および在庫分析データからこのブランドに属するすべてのレコードをフィルタリング
            const brandInventories = inventoryData.filter(inv => csCodesOfBrand.includes(inv['CS品番']?.value));
            const brandAnalyses = inventoryAnalysisData.filter(an => csCodesOfBrand.includes(an['CS品番']?.value));
            const sg = saleGoodsData.find(sg => sg['ブランド品番']?.value === brandCode);
            
            // 1. 各数値データの合計（合算）
            let periodSalesQty = 0;
            let periodSalesAmount = 0;
            csCodesOfBrand.forEach(cs => {
              periodSalesQty += salesSummaryByCs[cs] || 0;
              periodSalesAmount += salesAmountSummaryByCs[cs] || 0;
            });
            
            let totalInventoryQty = 0;
            brandInventories.forEach(inv => {
              totalInventoryQty += Number(inv['在庫数']?.value || 0);
            });

            // 在庫が1以上のブランド品番のみを出力（1未満はスキップ）
            if (totalInventoryQty < 1) {
              return null;
            }
            
            let totalSales30Days = 0;
            let totalSales7Days = 0;
            let totalFavoriteCount = 0;
            brandAnalyses.forEach(an => {
              totalSales30Days += Number(an['直近30日販売数']?.value || 0);
              totalSales7Days += Number(an['直近7日販売数']?.value || 0);
              totalFavoriteCount += Number(an['お気に入り登録数']?.value || 0);
            });
            
            // 2. テキスト・マスタデータの引き当てと補完
            // 親カテゴリ名
            const firstAnalysisWithParent = brandAnalyses.find(an => an['親カテゴリ']?.value || an['親カテゴリ名']?.value);
            const parentCategoryNameFromAnalysis = firstAnalysisWithParent?.['親カテゴリ']?.value || firstAnalysisWithParent?.['親カテゴリ名']?.value || '';
            
            const makerCode = (representativeProduct['メーカーID']?.value || representativeProduct['登録メーカーID']?.value || '')?.trim()?.replace(/\r?\n/g, '');
            const fallbackParentCategoryName = parentCategoryRecord?.['親カテゴリ名']?.value 
              || parentCategoryRecord?.['親カテゴリ']?.value 
              || parentCategoryRecord?.['親カテゴリー名']?.value 
              || parentCategoryRecord?.['メーカー名']?.value 
              || '';
            const parentCategoryName = parentCategoryNameFromAnalysis || fallbackParentCategoryName;
            
            // 子カテゴリ名
            const firstAnalysisWithChild = brandAnalyses.find(an => an['子カテゴリ']?.value || an['子カテゴリ名']?.value);
            const childCategoryNameFromAnalysis = firstAnalysisWithChild?.['子カテゴリ']?.value || firstAnalysisWithChild?.['子カテゴリ名']?.value || '';
            
            let resolvedChildCategoryName = childCategoryNameFromAnalysis;
            if (!resolvedChildCategoryName) {
              for (const prod of brandProducts) {
                const prodChildCategory = prod['子カテゴリ']?.value || '';
                const prodChildCategoryCode = prod['子カテゴリコード']?.value || '';
                
                if (prodChildCategory && isNaN(Number(prodChildCategory))) {
                  resolvedChildCategoryName = prodChildCategory;
                  break;
                } else if (prodChildCategoryCode) {
                  const subCatRecord = subCategoryMapById[prodChildCategoryCode];
                  if (subCatRecord?.['子カテゴリ名']?.value) {
                    resolvedChildCategoryName = subCatRecord['子カテゴリ名'].value;
                    break;
                  }
                } else if (prodChildCategory) {
                  resolvedChildCategoryName = prodChildCategory;
                  break;
                }
              }
            }
            const childCategoryName = resolvedChildCategoryName;
            
            // 主性別
            const firstAnalysisWithGender = brandAnalyses.find(an => an['主性別']?.value);
            const gender = firstAnalysisWithGender?.['主性別']?.value || representativeProduct['主性別']?.value || '';
            
            // 商品タイプ (親) / (子)
            const firstAnalysisWithParentType = brandAnalyses.find(an => an['親商品タイプ']?.value);
            let parentProductType = firstAnalysisWithParentType?.['親商品タイプ']?.value || '';
            
            const firstAnalysisWithChildType = brandAnalyses.find(an => an['子商品タイプ']?.value);
            let childProductType = firstAnalysisWithChildType?.['子商品タイプ']?.value || '';

            // 商品タイプが空の場合、商品マスタの商品タイプコードを元にZOZO商品タイプマスタから補完
            if (!parentProductType || !childProductType) {
              const firstProductWithType = brandProducts.find(p => p['商品タイプコード']?.value);
              const typeCode = firstProductWithType?.['商品タイプコード']?.value;
              if (typeCode) {
                const zozoType = zozoTypeMapById[String(typeCode).trim()];
                if (zozoType) {
                  if (!parentProductType) parentProductType = zozoType.parentType;
                  if (!childProductType) childProductType = zozoType.childType;
                }
              }
            }
            
            // 品名
            const firstAnalysisWithName = brandAnalyses.find(an => an['商品名']?.value);
            const firstInventoryWithName = brandInventories.find(inv => inv['商品名']?.value);
            const productName = firstAnalysisWithName?.['商品名']?.value || firstInventoryWithName?.['商品名']?.value || representativeProduct['商品名']?.value || '';
            
            // プロパー価格
            const firstAnalysisWithProperPrice = brandAnalyses.find(an => an['プロパー価格_税抜_']?.value);
            const firstInventoryWithProperPrice = brandInventories.find(inv => inv['プロパー価格_税抜_']?.value);
            const firstProductWithProperPrice = brandProducts.find(p => p['商品単価']?.value);
            
            const properPrice = firstAnalysisWithProperPrice?.['プロパー価格_税抜_']?.value 
              || firstInventoryWithProperPrice?.['プロパー価格_税抜_']?.value 
              || firstProductWithProperPrice?.['商品単価']?.value 
              || '';
            
            // 直近セール価格（税抜き）: salegoods の「変更後セール価格_税抜_」
            const recentSalePriceExcludingTax = sg?.['変更後セール価格_税抜_']?.value || '';
            // 直近セールオフ率: salegoods の「オフ率」
            const recentSaleOffRate = sg?.['オフ率']?.value || '';
            
            // 販売開始日: salegoods の「最新販売開始日」（ISO 8601形式 2026-05-17T06:37:00Z を YYYY-MM-DD 表記に変換）
            const rawSaleStartDate = sg?.['最新販売開始日']?.value || '';
            let saleStartDateValue = '';
            if (rawSaleStartDate) {
              const d = new Date(rawSaleStartDate);
              if (!isNaN(d.getTime())) {
                const yyyy = d.getFullYear();
                const mm = String(d.getMonth() + 1).padStart(2, '0');
                const dd = String(d.getDate()).padStart(2, '0');
                saleStartDateValue = `${yyyy}-${mm}-${dd}`;
              } else {
                saleStartDateValue = rawSaleStartDate;
              }
            }
            
            return [
              parentCategoryName,                                           // A: 親カテゴリ
              childCategoryName,                                            // B: 子カテゴリ
              brandCode,                                                    // C: ブランド品番
              gender,                                                       // D: 主性別
              parentProductType,                                            // E: (親)
              childProductType,                                             // F: (子)
              productName,                                                  // G: 品名
              properPrice,                                                  // H: プロパー価格（税抜）
              periodSalesQty,                                               // I: 受注数
              periodSalesAmount,                                            // J: 売上額
              recentSalePriceExcludingTax,                                  // K: 直近セール価格（税抜き）
              recentSaleOffRate,                                            // L: 直近セールオフ率
              '',                                                           // M: 直近セール価格（税込み）(記入しない)
              '',                                                           // N: オフ率から設定(↓) (記入しない)
              '',                                                           // O: 税込価格から設定(↓) (記入しない)
              '',                                                           // P: オフ率 (記入しない)
              '',                                                           // Q: セール価格（税込み） (記入しない)
              '',                                                           // R: セール価格（税抜き） (記入しない)
              saleStartDateValue,                                           // S: 販売開始日
              Number(totalInventoryQty) || 0,                               // T: 在庫
              Number(totalSales30Days) || 0,                                // U: 直近30日販売数
              Number(totalSales7Days) || 0,                                 // V: 直近7日販売数
              Number(totalFavoriteCount) || 0                               // W: 直近30日お気に入り登録数
            ];
          }).filter(Boolean);
          
          // 在庫が1以上のブランド品番が1件もない場合はファイル出力をスキップ
          if (excelRows.length === 0) {
            console.log(`[スキップ] ${progressLabel}: 在庫が1以上のブランド品番が存在しないため、ファイル出力をスキップします。`);
            isSuccess = true;
            continue;
          }
          
          // (4) テンプレートSpreadsheetのコピーを作成または既存のものを再利用して流し込み
          const startDateStr = saleStartDate ? saleStartDate.replace(/-/g, '') : '';
          const endDateStr = saleEndDate ? saleEndDate.replace(/-/g, '') : '';
          const periodStr = (startDateStr && endDateStr) ? `${startDateStr}～${endDateStr}` : formattedDateNoSlash;
          
          let displayName = '';
          const companyName = makerRecord?.['メーカー名']?.value || parentCategoryRecord?.['メーカー名']?.value || activeProducts[0]['メーカー名']?.value || makerName;
          if (meta.outputUnit === 'メーカー') {
            displayName = `${companyName}様`;
          } else {
            // 親カテゴリ単位の場合は「会社名_ブランド名様」とする
            const pName = parentCategoryRecord?.['親カテゴリ名']?.value || parentCategoryRecord?.['親カテゴリ']?.value || parentCategoryRecord?.['親カテゴリー名']?.value || '';
            displayName = pName ? `${companyName}_${pName}様` : `${companyName}様`;
          }
          const fileName = `${meta.makerId}_${displayName}_セール価格設定_${periodStr}`;
          
          if (!copiedSpreadsheet) {
            updateStatus('テンプレートのコピーを作成中... ' + progressLabel);
            copiedSpreadsheet = await copyDriveFileSale(
              accessToken,
              templateSpreadsheetId,
              fileName,
              dateFolderId,
              'application/vnd.google-apps.spreadsheet'
            );
          }
          
          // A〜L列 (index 0〜11) のデータと S〜W列 (index 18〜22) のデータに分割して書き込む（M〜R列の計算式を保護するため）
          const leftSideRows = excelRows.map(row => row.slice(0, 12));
          const rightSideRows = excelRows.map(row => row.slice(18, 23));
          
          updateStatus('データを書き込み中... ' + progressLabel);
          await updateSpreadsheetValuesBatch(accessToken, copiedSpreadsheet.id, [
            {
              range: 'A3',
              values: leftSideRows
            },
            {
              range: 'S3',
              values: rightSideRows
            }
          ]);
          
          // (5) Excelファイルに変換してKintoneへアップロード・追加
          updateStatus('Excel形式に変換中... ' + progressLabel);
          const exportedXlsxBlob = await exportSpreadsheetAsXlsxSale(accessToken, copiedSpreadsheet.id);
          const exportedXlsxFileName = `${fileName}.xlsx`;
          
          updateStatus('kintoneへアップロード中... ' + progressLabel);
          const exportedXlsxFileKey = await uploadFileToKintone(exportedXlsxBlob, exportedXlsxFileName);
          
          // 各種宛先情報の抽出（連絡先テーブルから宛先種別に応じて TO, CC を抽出）
          const targetRecord = meta.outputUnit === 'メーカー' ? makerRecord : parentCategoryRecord;
          
          let nameTo = [];
          let mailTo = [];
          let nameCc = [];
          let mailCc = [];
          
          if (targetRecord && targetRecord['連絡先']?.value && targetRecord['連絡先'].value.length > 0) {
            const contacts = targetRecord['連絡先'].value;
            
            // TO
            contacts.filter(item => item.value['宛先種別']?.value?.includes("セール（TO）")).forEach(item => {
              nameTo.push(item.value['担当者名']?.value || '');
              mailTo.push(item.value['メールアドレス']?.value || '');
            });
            // CC
            contacts.filter(item => item.value['宛先種別']?.value?.includes("セール（CC）")).forEach(item => {
              nameCc.push(item.value['担当者名']?.value || '');
              mailCc.push(item.value['メールアドレス']?.value || '');
            });
          }
          
          const recipientToName = nameTo.filter(Boolean).join(',');
          const recipientToEmail = mailTo.filter(Boolean).join(',');
          const recipientCcName = nameCc.filter(Boolean).join(',');
          const recipientCcEmail = mailCc.filter(Boolean).join(',');

          // FA様担当者の取得（メーカーマスタの「FUN担当者ID」から取得し、ユーザー選択型の形式 [{ code: "..." }] に変換）
          let faUser = null;
          const rawFunUser = makerRecord?.['FUN担当者ID']?.value;
          if (Array.isArray(rawFunUser) && rawFunUser.length > 0) {
            faUser = rawFunUser.map(u => ({ code: (typeof u === 'object' && u.code) ? u.code : String(u) })).filter(u => u.code);
            if (faUser.length === 0) faUser = null;
          } else if (typeof rawFunUser === 'string' && rawFunUser.trim()) {
            faUser = [{ code: rawFunUser.trim() }];
          } else if (typeof rawFunUser === 'number') {
            faUser = [{ code: String(rawFunUser) }];
          }

          const appId = config['セール返答表'].id;
          
          // 値が未定義・空文字の場合はすべて null に統一するヘルパー関数
          const clean = (val) => (val === undefined || val === null || val === "") ? null : val;

          const rawRecord = {
            "セール設定レコード番号": clean(recordId),
            "セール開始日": clean(saleStartDate),
            "セール終了日": clean(saleEndDate),
            "メーカー名": clean(makerName),
            "メーカーコード": clean(meta.makerId),
            "ブランド名": clean(brandName),
            "親カテゴリーコード": clean(finalParentCategoryCode),
            "セール価格設定リスト": exportedXlsxFileKey ? [{ fileKey: exportedXlsxFileKey }] : null,
            "FA様担当者": faUser,
            "ブランド連絡先担当者名": clean(recipientToName),
            "ブランド連絡先メールアドレス": clean(recipientToEmail),
            "ブランド連絡先担当者名_CC_": clean(recipientCcName),
            "ブランド連絡先メールアドレス_CC_": clean(recipientCcEmail)
          };

          // 値が null のプロパティはリクエストの record から除外する (不正なJSON文字列エラーを確実に防止するため)
          const recordData = {};
          Object.keys(rawRecord).forEach(key => {
            if (rawRecord[key] !== null) {
              recordData[key] = { value: rawRecord[key] };
            }
          });

          // 個別追加の代わりに、一括登録用の配列に追加する
          recordsToRegister.push(recordData);
          
          isSuccess = true;
          successCount++;
          console.log(`--- [セール設定処理準備完了] ${progressLabel} ---`);
          
          // Google API への連続アクセス（レート制限）を和らげるためのウェイト（1.0秒）
          await new Promise(resolve => setTimeout(resolve, 1000));
          
        } catch (makerError) {
          console.error(`【デバッグエラー】メーカー/ブランド「${makerName}」の試行 ${attempt} 回目でエラーが発生しました。`, makerError);
          if (attempt < maxRetries) {
            const retryWait = attempt * 3000; // 3秒、6秒、9秒...と段階的に長く待つ（指数バックオフ）
            updateStatus(`通信エラーのため、${retryWait / 1000}秒後に再試行します... ` + progressLabel);
            await new Promise(resolve => setTimeout(resolve, retryWait));
          } else {
            failCount++;
            failedManufacturers.push(makerName);
            console.error(`【エラー発生】メーカー/ブランド「${makerName}」の処理中にエラーが発生しました。スキップします。`, makerError);
          }
        }
      }
    }
    
    // ----------------------------------------------------
    // 4. kintone へのレコード一括登録
    // ----------------------------------------------------
    if (recordsToRegister.length > 0) {
      updateStatus(`kintone へレコードを一括登録中（全 ${recordsToRegister.length} 件）...`);
      const appId = config['セール返答表'].id;
      const registeredIds = await registerRecordsInBulk(appId, recordsToRegister);
      console.log(`【一括登録完了】レコードID群: ${registeredIds.join(', ')}`);
    }
    
    if (failCount > 0) {
      updateStatus(`処理完了（成功: ${successCount}件, 失敗: ${failCount}件）。詳細はコンソールログを確認してください。`);
      console.warn('処理に失敗したメーカー/ブランド一覧:', failedManufacturers);
    } else {
      updateStatus('すべてのセール設定リスト出力が完了しました！');
    }
    
  } catch (error) {
    updateStatus('エラーが発生しました。');
    console.error('処理中に失敗しました。', error);
    let detailMsg = error.message || '不明なエラー';
    if (error.results) {
      const details = Object.entries(error.results).map(([key, val]) => {
        return `${key}: ${val.error || JSON.stringify(val)}`;
      }).join('\n');
      detailMsg += '\n【詳細】\n' + details;
    } else if (error.errors) {
      const details = Object.entries(error.errors).map(([key, val]) => {
        return `${key}: ${val.messages ? val.messages.join(', ') : JSON.stringify(val)}`;
      }).join('\n');
      detailMsg += '\n【詳細】\n' + details;
    }
    throw new Error(detailMsg);
  }
}

/**
 * Google Sheets API の batchUpdate を呼び出して、複数範囲を1回で一括更新する
 */
async function updateSpreadsheetValuesBatch(accessToken, spreadsheetId, valueRanges) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`;
  const response = await fetchWithRetryOnce(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      valueInputOption: 'USER_ENTERED',
      data: valueRanges
    })
  });
  
  const data = await response.json();
  if (!response.ok) {
    const message = data && data.error && data.error.message ? data.error.message : JSON.stringify(data);
    throw new Error(`Google Sheets batchUpdate Error: ${message}`);
  }
  return data;
}

/**
 * kintone の複数レコード一括登録APIを、100件ずつに分割しながら実行する
 */
async function registerRecordsInBulk(appId, recordsList) {
  const chunkSize = 100;
  const ids = [];
  for (let i = 0; i < recordsList.length; i += chunkSize) {
    const chunk = recordsList.slice(i, i + chunkSize);
    const requestBody = {
      app: appId,
      records: chunk
    };
    console.log(`【一括登録】${chunk.length}件のレコードを送信中...`);
    const resp = await kintone.api(kintone.api.url('/k/v1/records.json', true), 'POST', requestBody);
    ids.push(...resp.ids);
  }
  return ids;
}

/**
 * 対象コード（品番など）の配列を一定件数（500件）ずつに分割し、kintoneの「in」クエリで小分けに取得・マージする
 */
async function fetchRecordsInChunks(client, appId, codesList, fieldName, baseCondition = '') {
  if (!codesList || codesList.length === 0) return [];
  
  const chunkSize = 500;
  const allRecords = [];
  
  for (let i = 0; i < codesList.length; i += chunkSize) {
    const chunk = codesList.slice(i, i + chunkSize);
    let query = `${fieldName} in (${chunk.map(v => '"' + String(v).replace(/"/g, '\\"') + '"').join(',')})`;
    if (baseCondition) {
      query = `(${baseCondition}) and (${query})`;
    }
    
    console.log(`【分割取得】appId: ${appId}, ${i}件目からの ${chunk.length}件を取得中...`);
    const records = await client.record.getAllRecords({
      app: appId,
      condition: query
    });
    allRecords.push(...records);
  }
  
  return allRecords;
}

// =========================================================================
// セール設定専用の Google API 通信強化関数群（他機能への影響を避けるため main.js 内に定義）
// =========================================================================

async function fetchWithRetryOnceSale(url, options) {
  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = await fetch(url, options);
      if (response.ok) {
        return response;
      }
      if (response.status === 429 || (response.status >= 500 && response.status < 600)) {
        if (attempt === maxAttempts) return response;
        const wait = attempt * 1500;
        console.warn(`[Google API 一時的エラー ${response.status}] ${wait / 1000}秒後に再試行します... (試行: ${attempt}/${maxAttempts})`);
        await new Promise(resolve => setTimeout(resolve, wait));
        continue;
      }
      return response;
    } catch (error) {
      if (attempt === maxAttempts) {
        throw error;
      }
      const wait = attempt * 1500;
      console.warn(`[通信エラー Failed to fetch] ${wait / 1000}秒後に再試行します... (試行: ${attempt}/${maxAttempts})`, error);
      await new Promise(resolve => setTimeout(resolve, wait));
    }
  }
}

async function findDriveFileSale(accessToken, parentFolderId, fileName) {
  if (!fileName) {
    throw new Error('File name is required.');
  }
  const queryParts = [
    "mimeType != 'application/vnd.google-apps.folder'",
    'trashed = false',
    `name = '${String(fileName).replace(/'/g, "\\'")}'`
  ];
  if (parentFolderId) {
    queryParts.push(`'${String(parentFolderId).replace(/'/g, "\\'")}' in parents`);
  }
  const params = new URLSearchParams({
    q: queryParts.join(' and '),
    fields: 'files(id,name,webViewLink,parents,mimeType)',
    pageSize: '1',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true'
  });
  const response = await fetchWithRetryOnceSale(`https://www.googleapis.com/drive/v3/files?${params.toString()}`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
  const data = await response.json();
  if (!response.ok) {
    const message = data && data.error && data.error.message ? data.error.message : JSON.stringify(data);
    throw new Error('Drive File Search Error: ' + message);
  }
  return Array.isArray(data.files) && data.files.length > 0 ? data.files[0] : null;
}

async function deleteDriveItemSale(accessToken, fileId) {
  if (!fileId) {
    throw new Error('File ID is required.');
  }
  const params = new URLSearchParams({
    supportsAllDrives: 'true'
  });
  const response = await fetchWithRetryOnceSale(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?${params.toString()}`, {
    method: 'DELETE',
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error('Drive Item Delete Error: ' + errorText);
  }
}

async function copyDriveFileSale(accessToken, fileId, newName, parentFolderId, targetMimeType = '') {
  const existingFile = await findDriveFileSale(accessToken, parentFolderId, newName);
  if (existingFile) {
    await deleteDriveItemSale(accessToken, existingFile.id);
  }
  const payload = {
    name: newName
  };
  if (parentFolderId) payload.parents = [parentFolderId];
  if (targetMimeType) payload.mimeType = targetMimeType;

  const response = await fetchWithRetryOnceSale(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/copy?fields=id,name,webViewLink,parents,mimeType`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok) {
    const message = data && data.error && data.error.message ? data.error.message : JSON.stringify(data);
    throw new Error('Drive File Copy Error: ' + message);
  }
  return data;
}

async function exportSpreadsheetAsXlsxSale(accessToken, spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error('Spreadsheet ID is required.');
  }
  const response = await fetchWithRetryOnceSale(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(spreadsheetId)}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const message = errorData && errorData.error && errorData.error.message ? errorData.error.message : response.statusText;
    throw new Error('Spreadsheet Export Error: ' + message);
  }
  return await response.blob();
}
