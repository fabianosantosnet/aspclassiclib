<%
dim nl 
nl = chr(13) 

'num_docformat 1  ( XHTML 1.0 Strict ), 2  ( XHTML 1.0 Transitional ), 3  ( XHTML 1.0 Frameset ), 4 (HTML 5), 5  ( HTML 4.01 Strict ), 6 ( HTML 4.01 Transitional ), 7  ( HTML 4.01 Frameset )
sub html_doctype(docformat)

    if docformat = 1 then %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"><% 
    elseif docformat = 2 then %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><%
    elseif docformat = 3 then %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd"><%

    elseif docformat = 4 then %><!DOCTYPE html><%

    elseif docformat = 5  then %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd"><%
    elseif docformat = 6 then %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd"><%
    elseif docformat = 7 then %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd"><%
    end if
end sub

sub html()
    response.write nl & "<html xmlns='http://www.w3.org/1999/xhtml'>" 
end sub

sub html_header(dados)
    response.write( nl & nl & "<head>"& nl & dados & nl & "</head>" & nl )
end sub

function html_title(title)
    html_title = "<title>" &  title & "</title>"
end function

'Info em https://developer.mozilla.org/pt/Utilizando_meta_tags
function html_meta_author(person, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    html_meta_author = "<meta name='author' content='" & person & "'" & ch &">"    
    
end function

function html_meta_language(typeLang, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    if typeLang = 1 then
        html_meta_language= "<meta http-equiv='content-language' content='pt-br'" & ch &">" & nl & "<meta name='language' content='pt-br'"& ch &">"
    elseif typeLang = 2 then
        html_meta_language= "<meta http-equiv='content-language' content='en-us'" & ch &">" & nl & "<meta name='language' content='en-us'" & ch &">"
    end if
    
end function

function html_meta_copyright(dono, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    html_meta_copyright = "<meta name='copyright' content='© " & dono & "'" & ch &">"
    
end function

function html_meta_charset(type_charset, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    if type_charset = 1 then
        html_meta_charset = "<meta http-equiv='content-type' content='text/html; charset=iso-8859-1'"& ch &">"
    elseif type_charset = 2 then
        html_meta_charset = "<meta http-equiv='content-type' content='text/html; charset=utf-8'"& ch &">"
    end if
end function

function html_meta_robots(type_robots, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    if type_robots = 1 then
        html_meta_robots = "<meta name='robots' content='all'" & ch &">"
    elseif type_robots = 2 then
        html_meta_robots = "<meta name='robots' content='index,follow'" & ch &">"
    elseif type_robots = 3 then
        html_meta_robots = "<meta name='robots' content='noindex,follow'" & ch &">"          
    elseif type_robots = 4 then
        html_meta_robots = "<meta name='robots' content='index,nofollow'" & ch &">"                  
    elseif type_robots = 5 then
        html_meta_robots = "<meta name='robots' content='noindex,nofollow'" & ch &">"             
    end if
end function

function html_meta_description(desc, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if   
            
     html_meta_description = "<meta name='description' content='" & desc & "'" & ch & ">"
    
end function


function html_meta_keyword(key, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if   
            
     html_meta_keyword = "<meta name='keywords' content='" & key & "'" & ch & ">"
    
end function

function html_meta_cache(typeCache,isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    if typeCache = 1 then
        html_meta_cache = "<meta http-equiv='pragma' content='no-cache'"& ch &">" & nl & "<meta http-equiv='cache-control' content='no-cache'"& ch &">"  
     elseif typeCache = 2 then
        html_meta_cache = "<meta http-equiv='cache-control' content='no-store'"& ch &">" 
     elseif typeCache = 3 then
        html_meta_cache = "<meta http-equiv='cache-control' content='public'"& ch &">"         
     elseif typeCache = 4 then
        html_meta_cache = "<meta http-equiv='cache-control' content='private'"& ch &">" 
     end if     
     
    end function
 
    'No internet explorer elimina aquela pequena barra de opções que aparece ao passarmos o mouse por cima de uma imagem
    function html_meta_imagetoolbar(isXHTML)
        dim ch
    
        if isXHTML = true then
            ch = " /" 
        end if   
 
        html_meta_imagetoolbar ="<meta http-equiv='imagetoolbar' content='no'" & ch & ">" 
        
     end function
        
    function html_meta_refresh(timesec,url, isXHTML)
        dim ch
    
        if isXHTML = true then
            ch = " /" 
        end if
       
        html_meta_refresh = "<meta http-equiv='refresh' content='" & timesec & ";url=" & url & "'" & ch &">"
        
    end function
    
    function html_body(text)
        response.write "<body" & text & ">"
    end function

    function html_footer()
        response.write "</body>" & nl & "</html>"
    end function
    
'Format date Mon, 22 jul 2006 11:12:01 GMT    
function html_meta_expires(date, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    html_meta_expires = "<meta http-equiv='expires' content='" & date & "'" & ch &">"
    
end function

function html_meta_generator(generator, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    html_meta_generator = "<meta name='generator' content='" & generator & "'" & ch &">"
    
end function

'set revTime value "7 days" or "1 month"
function html_meta_revisit(revTime, isXHTML)
    dim ch
    
    if isXHTML = true then
        ch = " /" 
    end if
    
    html_meta_revisit = "<meta name='revisit-after' content='" & revTime & "'" & ch &">"
    
end function
    
'Copy loop at Criarweb
sub all_serverVariable()
    for each i in Request.ServerVariables
        response.write i & " : " & Request.ServerVariables(i) & "<br>"   
    Next
end sub

%>


