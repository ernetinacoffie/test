const axios = require('axios');
const cheerio = require('cheerio')
const path = require('path');
const fs = require('fs');
const nodeXlsx = require('node-xlsx')
const SocksProxyAgent = require('socks-proxy-agent')
const moment = require('moment')
const config = require('./config.js')
const UserAgent = require('user-agents');
const puppeteer = require('puppeteer');
const { plugin } = require('puppeteer-with-fingerprints');

console.log(config)

const httpsAgent = new SocksProxyAgent.SocksProxyAgent('socks5://127.0.0.1:7890')

var address = config.address
var browser = null
var page_num = 1
var page_index = 1

const getPage = ()=>{
	return new Promise((reslove,reject)=>{
		try{
			var page = fs.readFileSync('./pageInfo.json','utf-8')
			page = JSON.parse(page)
		}catch{
			reject('页面文件读取错误')
			return 
		}
		page.forEach(item=>{
			if(item.page == address){
				page_num = item.page_num
				page_index = item.index
			}
		})
		reslove('end')
	})
}

const start = async ()=>{
	var getPageState = await getPage().catch(error=>{
	 	console.log(error)
	 })
	if(getPageState != 'end') return
	visitPage()
}
start()

const visitPage = async ()=>{
	console.log('当前进行第：'+page_num+'页')
	var webAddress = address+'/'+page_num
	var pagelist = await axios(webAddress,{
		httpsAgent
	}).catch(error=>{
		console.log('页面获取错误,正在重试：'+'\n'+webAddress)
		setTimeout(function(){
			visitPage()
		},3000)
	})
	let $ = cheerio.load(pagelist.data);
	var companyLinkArr = []
	$('td a').each((index,item)=>{
		if($(item).text() == 'Check company details'){
			let href = item.attribs.href
			companyLinkArr.push('http://www.datalog.co.uk'+href)
		}
	})
    let status = await getCompany(companyLinkArr)
	if(status != 'end') return console.log('网络错误')
	page_num++
	page_index = 1
	await writConfig()
	if(page_num>500) return console.log('结束')
	setTimeout(function(){
		visitPage()
	},3000)
}

const getCompany = async (linkArr)=>{
	return new Promise( async (reslove,reject)=>{
		if(linkArr.length){
			let link = linkArr.shift()
			var company = {
					name: null,
					id: null,
					country: null,
					type: null,
					status: null,
					vat:null,
					page:address,
					page_num:page_num,
					index:page_index,
					careateTime:moment().format('YYYY-MM-DD HH:mm:ss')
				}
			company.name = decodeURIComponent(link.split('/').pop()).replaceAll('+',' ')
			
			await openChrome()
			openPage(link).then(async page=>{
				let body = await getElement(page)
				let $ = cheerio.load(body);
				$('#filing').parent().find('tr').each((index,item)=>{
					if($(item).children('th').text() == 'Company ID Number'){
						company.id = $(item).children('td').text().trim()
					}
					if($(item).children('th').text() == 'Origin Country'){
						company.country = $(item).children('td').text().trim()
					}
					if($(item).children('th').text() == 'Type'){
						company.type = $(item).children('td').text().trim()
					}
					if($(item).children('th').text() == 'CompanyStatus'){
						company.status = $(item).children('td').text().trim()
					}
					if($(item).children('th').text() == 'VAT Number /Sales tax ID'){
						company.vat = $(item).children('td').text().trim()
					}
				})
				writeFile(company)
				setTimeout(function(){
					page.close()
				},2000)
			}).catch(error=>{
				company.type = '获取信息失败'
				writeFile(company)
				console.log('获取信息失败'+'\n'+link)
			}).finally(()=>{
				setTimeout(function(){
					reslove(getCompany(linkArr))
				},3000)
			})
		}else{
			reslove('end')
		}
	})
}

const openChrome = async ()=>{
	// const fingerprint = await plugin.fetch('', {
	//   tags: ['Microsoft Windows', 'Chrome'],
	// });
	// plugin.useProxy('socks5://127.0.0.1:7890',{
	// 	// Change browser timezone according to proxy:
	// 	changeTimezone: true,
	// 	// Replace browser geolocation according to proxy:
	// 	changeGeolocation: true,
	// }).useFingerprint(fingerprint);
	// const browser = await plugin.launch();
	// pageInstance = await browser.newPage();
	// await pageInstance.goto('https://browserleaks.com/canvas', { waitUntil: 'networkidle0' });
	// console.log(pageInstance)
	return new Promise(async (reslove)=>{
		if(browser) return reslove(browser)
		browser = await puppeteer.launch({
		 	ignoreHTTPSErrors:true,
		 	executablePath:'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
		 	headless: false,
		});
		reslove(browser)
	})
}

const openPage = async (url)=>{
	return new Promise(async (reslove,reject)=>{
		try{
			var pageInstance = await browser.newPage();
			var userAgent  = new UserAgent()
			pageInstance.setUserAgent(userAgent.toString());
			pageInstance.setViewport(setViewport())
			await pageInstance.goto(url,{
				timeout:0
			})
			reslove(pageInstance)
		}catch(error){
			reject(error)
		}
	})
}

const getElement = async (page)=>{
	return new Promise(async (reslove,reject)=>{
		let filing = page.$('#filing',e=>e.outerHTML)
		if(!filing){
			console.log('请手动进行人机验证，或错误处理')
			setTimeout(function(){
				reslove(getElement(page))
			},1000)
		}else{
			let body = await page.$eval('body',e=>e.outerHTML)
			reslove(body)
		}
	})
}

const setViewport = ()=>{
	var w = 800
	var h = 600
	w = Math.floor(Math.random()*1000)+600
	h = Math.floor(((Math.random()*30)+50)/100)
	return {
		width:w,
 		height:h
	}
}
const writeFile = (company)=>{
	console.log('完成第'+page_num+'页，第'+page_index+'条：'+company.name)
	page_index++
	var ex1 = nodeXlsx.parse("./companyList.xlsx")
	excel_content = ex1[0].data

	var str = [company.name,
			company.id,
			company.vat,
			company.country,
			company.type,
			company.status,
			company.careateTime,
			company.page_num,
			company.index,
			company.page]
	// 将新内容添加到工作表
	excel_content.push(str);

	// 创建新的工作簿并添加更新后的工作表
	const updatedWorkbook = nodeXlsx.build([{ name: ex1[0].name, data: excel_content }]);

	// 保存更新后的 XLSX 文件
	fs.writeFileSync("./companyList.xlsx", updatedWorkbook);
}


const writConfig = ()=>{
	return new Promise((reslove,reject)=>{
		try{
			var page = fs.readFileSync('./pageInfo.json','utf-8')
			page = JSON.parse(page)
		
			let s = true
			page.forEach(item=>{
				if(item.page == address){
					s = false
					item.page_num = page_num
					item.index = page_index 
				}
			})
			if(s){
				page.push({
					page:address,
					page_num:page_num,
					index:1
				})
			}
			fs.writeFile('./pageInfo.json', JSON.stringify(page),'utf8', err => {
			 if (err) {
			    console.error(err);
			  }
			  // file written successfully
			})
		}catch (error){
			console.log(error)
			console.log('页面文件读取错误')
		}
		reslove('end')
	})
}