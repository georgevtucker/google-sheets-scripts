function getPageTitle(url){
  var result = UrlFetchApp.fetch( url );
  var wholePage = result.getContentText(); 
  var scrap = wholePage.match( /<title>(.*?)<\/title>/ ); 
  var title = scrap[1];
  return scrap;
}  
