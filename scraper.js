const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const path = require('path');

async function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function getProductInfo(page, url) {
  try {
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 20000 });
    
    // Wait for any of the elements to load
    await Promise.race([
      page.waitForSelector('.seller-information a', { timeout: 4000 }),
      page.waitForSelector('.title-content-list .brand-link a', { timeout: 2000 }),
      page.waitForSelector('.description-card-module_description-card_m9PqC', { timeout: 5000 })
    ]);
    
    // Extract all information in one evaluation
    const result = await page.evaluate(() => {
      const sellerElement = document.querySelector('.seller-information a');
      const brandElement = document.querySelector('.title-content-list .brand-link a');
      const descriptionElement = document.querySelector('.product-description.product-description-module_product-description_3bMdX');
      
      // Extract additional sellers count
      const offersElement = document.querySelector('.more-buying-choices-module_offer_34xYl');
      let additionalSellers = 'No additional sellers';
      
      if (offersElement) {
        const offerText = offersElement.textContent.trim();
        const match = offerText.match(/(\d+)\s+offer/);
        if (match) {
          additionalSellers = match[1];
        }
      }

      // Extract list price (price before markdown)
      let listPrice = 'N/A';
      const listPriceElement = document.querySelector('.buybox-offer-module_list-price_2GEsn .currency');
      if (listPriceElement) {
        listPrice = listPriceElement.textContent.trim();
      }
      
      // Clean up description text
      const cleanDescription = (text) => {
        if (!text) return '';
        return text
          .replace(/\s+/g, ' ')
          .replace(/\n\s*\n/g, '\n')
          .trim();
      };
      
      return {
        seller: sellerElement ? sellerElement.textContent.trim() : 'No seller name found',
        brand: brandElement ? brandElement.textContent.trim() : 'No brand found',
        description: descriptionElement ? cleanDescription(descriptionElement.innerText) : 'No description found',
        additionalSellers: additionalSellers,
        listPrice: listPrice
      };
    });

    return result;
  } catch (err) {
    console.error('Error fetching product info:', err.message);
    return {
      seller: 'Error fetching seller',
      brand: 'Error fetching brand',
      description: 'Error fetching description',
      additionalSellers: 'Error fetching additional sellers',
      listPrice: 'Error fetching list price'
    };
  }
}

async function processBatch(products, startIdx, batchSize, browser) {
  const batchPromises = [];
  const endIdx = Math.min(startIdx + batchSize, products.length);

  for (let i = startIdx; i < endIdx; i++) {
    const product = products[i];
    batchPromises.push((async () => {
      const page = await browser.newPage();
      try {
        const { seller, brand, description, additionalSellers, listPrice } = await getProductInfo(page, product['Product URL']);
        return {
          ...product,
          Seller: seller,
          Brand: brand,
          Description: description,
          'Additional Sellers': additionalSellers,
          'List Price': listPrice
        };
      } finally {
        await page.close();
      }
    })());
  }

  return Promise.all(batchPromises);
}

async function getAllProducts(url) {
  const browser = await puppeteer.launch({ 
    headless: false,
    defaultViewport: null,
    args: ['--start-maximized']
  });
  
  const mainPage = await browser.newPage();

  try {
    await mainPage.goto(url, { waitUntil: 'networkidle0' });
    await mainPage.waitForSelector('article.product-card-module_product-card_fdqa8', { timeout: 10000 });

    let previousProductCount = 0;
    let attemptCount = 0;
    const maxAttempts = 5;

    console.log('Starting to load all products...');
    
    while (true) {
      const currentProductCount = await mainPage.evaluate(() => 
        document.querySelectorAll('article.product-card-module_product-card_fdqa8').length
      );
      
      console.log(`Current product count: ${currentProductCount}`);

      try {
        await mainPage.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
        await delay(1000);

        const buttonVisible = await mainPage.evaluate(() => {
          const button = document.querySelector('.search-listings-module_load-more_OwyvW');
          if (!button) return false;
          
          const rect = button.getBoundingClientRect();
          return rect.top >= 0 && rect.bottom <= window.innerHeight;
        });

        if (!buttonVisible) {
          console.log('Load more button not found or not visible');
          break;
        }

        await mainPage.click('.search-listings-module_load-more_OwyvW');
        console.log('Clicked load more button');
        await delay(2000);
        
        if (currentProductCount === previousProductCount) {
          attemptCount++;
          if (attemptCount >= maxAttempts) {
            console.log('No new products loaded after multiple attempts. Stopping.');
            break;
          }
        } else {
          attemptCount = 0;
        }

        previousProductCount = currentProductCount;

      } catch (err) {
        console.log('Error while loading more products:', err.message);
        break;
      }
    }

    console.log('Extracting product data...');
    
    const baseUrl = await mainPage.evaluate(() => window.location.origin);
    
    const products = await mainPage.evaluate((baseUrl) => {
      const items = [];
      
      document.querySelectorAll('article.product-card-module_product-card_fdqa8').forEach(article => {
        try {
          const linkElement = article.querySelector('a.product-card-module_link-underlay_3sfaA');
          const contentElement = article.querySelector('.grid-y.gap-2.product-card-module_product-card-content_1LFYj');
          
          if (linkElement && contentElement) {
            const href = linkElement.getAttribute('href');
            const title = contentElement.querySelector('.product-card-module_product-title_16xh8')?.textContent?.trim() || '';
            const price = contentElement.querySelector('.currency')?.textContent?.trim() || '';
            const image = contentElement.querySelector('img')?.src || '';
            
            // Extract list price from product card if available
            let listPrice = 'N/A';
            const listPriceElement = contentElement.querySelector('.product-card-price-module_list-price_om_3Y .currency');
            if (listPriceElement) {
              listPrice = listPriceElement.textContent.trim();
            }
            
            const productUrl = href ? `${baseUrl}${href}` : '';
            
            if (title && productUrl) {
              items.push({
                Title: title,
                Price: price,
                'List Price': listPrice,
                'Image URL': image,
                'Product URL': productUrl
              });
            }
          }
        } catch (error) {
          console.error('Error processing article:', error);
        }
      });
      
      return items;
    }, baseUrl);

    console.log('Getting seller information concurrently...');

    // Process products in batches of 5 concurrently
    const BATCH_SIZE = 1;
    const allProducts = [];
    
    for (let i = 0; i < products.length; i += BATCH_SIZE) {
      console.log(`Processing batch ${i/BATCH_SIZE + 1}/${Math.ceil(products.length/BATCH_SIZE)}`);
      const batchResults = await processBatch(products, i, BATCH_SIZE, browser);
      allProducts.push(...batchResults);
    }

    const uniqueProducts = [...new Map(allProducts.map(item => 
      [item.Title + item.Price, item]
    )).values()];

    console.log(`Total unique products found: ${uniqueProducts.length}`);
    return uniqueProducts;

  } catch (err) {
    console.error('Error during scraping:', err.message);
    return null;
  } finally {
    await browser.close();
  }
}

function exportToExcel(products, url) {
  try {
    const ws = XLSX.utils.json_to_sheet(products);

    const colWidths = [
      { wch: 50 }, // Title
      { wch: 15 }, // Price
      { wch: 15 }, // List Price
      { wch: 50 }, // Image URL
      { wch: 70 }, // Product URL
      { wch: 30 }, // Seller
      { wch: 20 }, // Brand
      { wch: 100 }, // Description
      { wch: 20 }  // Additional Sellers
    ];
    ws['!cols'] = colWidths;

    // Set row height for description column
    ws['!rows'] = Array(products.length + 1).fill({ hpt: 100 });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Products');

    const domain = new URL(url).hostname.replace(/[^a-z0-9]/gi, '_');
    const timestamp = new Date().toISOString().replace(/[^0-9]/g, '').slice(0, 14);
    const filename = `products_${domain}_${timestamp}.xlsx`;
    const filepath = path.join(process.cwd(), filename);

    // Configure sheet to wrap text in description column
    const descriptionCol = XLSX.utils.decode_col('G');
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      const cell = ws[XLSX.utils.encode_cell({r: R, c: descriptionCol})];
      if (cell) {
        if (!cell.s) cell.s = {};
        cell.s.alignment = { wrapText: true, vertical: 'top' };
      }
    }

    XLSX.writeFile(wb, filepath);
    console.log(`Excel file saved successfully: ${filepath}`);
    return filepath;
  } catch (err) {
    console.error('Error exporting to Excel:', err.message);
    return null;
  }
}

// Main execution
(async () => {
  const url = process.argv[2];
  if (!url) {
    console.log('Please provide a URL as an argument');
    process.exit(1);
  }

  console.log('Starting scraper...');
  const products = await getAllProducts(url);
  
  if (products && products.length > 0) {
    console.log(`Scraping completed. Found ${products.length} products.`);
    
    const excelFile = exportToExcel(products, url);
    if (excelFile) {
      console.log('\nScraping Summary:');
      console.log('----------------');
      console.log(`Total Products: ${products.length}`);
      console.log(`Excel File: ${excelFile}`);
    } else {
      console.log('Failed to create Excel file.');
    }
  } else {
    console.log('No products found or scraping failed.');
  }
})();