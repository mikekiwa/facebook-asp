<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #include file="fb_app.asp" -->
<%
session.lcid = 1046
session.timeOut = 1440
server.scriptTimeOut = 999999999
 
if (request.QueryString("logar") = "") then
    main
    function main
        dim strJSON
        dim URL
        dim sToken
        dim user
        dim loc
     
        set cookie = get_facebook_cookie( FACEBOOK_APP_ID, FACEBOOK_SECRET )
        if cookie.count > 0 then 
            response.write "Logado... Ok! <br/>"
            sToken = cookie("access_token")
            url = "https://graph.facebook.com/me?access_token=" & sToken
            strJSON = get_page_contents(URL) 
     
            set user = JSON.parse(strJSON)
            response.write cookie("access_token")
        else
            link = "http://www.facebook.com/dialog/oauth?client_id=ID_DA_SUA_APLICAÇÃO&redirect_uri="&server.URLEncode("http://apps.facebook.com/NOME_DA_MINHA_APP/autenticacao.asp?logar=ok")&"&scope=offline_access,user_location,email,publish_stream,user_birthday,read_friendlists"
            response.write("<script>top.location.href='" & link & "';</script>")
        end if
    end function
else
    const BASE_64_MAP_INIT ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    dim nl
    dim Base64EncMap(63)
    dim Base64DecMap(127)
     
    public sub initCodecs()
        nl = "<P>" & chr(13) & chr(10)
        dim max, idx
           max = len(BASE_64_MAP_INIT)
        for idx = 0 to max - 1
             Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
        next
        for idx = 0 to max - 1
             Base64DecMap(asc(Base64EncMap(idx))) = idx
        next
    end sub
     
    public function base64Encode(plain)
        if len(plain) = 0 then
             base64Encode = ""
             exit function
        end if
        dim ret, ndx, by3, first, second, third
        by3 = (len(plain) \ 3) * 3
        ndx = 1
        do while ndx <= by3
             first  = asc(mid(plain, ndx + 0, 1))
             second = asc(mid(plain, ndx + 1, 1))
             third  = asc(mid(plain, ndx + 2, 1))
             ret = ret & Base64EncMap((first \ 4) and 63)
             ret = ret & Base64EncMap(((first * 16) and 48) + ((second \ 16) and 15))
             ret = ret & Base64EncMap(((second * 4) and 60) + ((third \ 64) and 3))
             ret = ret & Base64EncMap(third and 63)
             ndx = ndx + 3
        loop
        if by3 < len(plain) then
             first  = asc(mid(plain, ndx + 0, 1))
             ret = ret & Base64EncMap((first \ 4) and 63)
             if (len(plain) MOD 3) = 2 then
                  second = asc(mid(plain, ndx+1, 1))
                  ret = ret & Base64EncMap(((first * 16) and 48) + ((second \ 16) and 15))
                  ret = ret & Base64EncMap(((second * 4) and 60))
             else
                  ret = ret & Base64EncMap((first * 16) and 48)
                  ret = ret & "="
             end if
             ret = ret & "="
        end if
        base64Encode = ret
    end function
     
    public function base64Decode(scrambled)
        if len(scrambled) = 0 then
             base64Decode = ""
             exit function
        end if
        dim realLen
        realLen = len(scrambled)
        do while mid(scrambled, realLen, 1) = "="
             realLen = realLen - 1
        loop
        dim ret, ndx, by4, first, second, third, fourth
        ret = ""
        by4 = (realLen \ 4) * 4
        ndx = 1
        do while ndx <= by4
             first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
             second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
             third  = Base64DecMap(asc(mid(scrambled, ndx+2, 1)))
             fourth = Base64DecMap(asc(mid(scrambled, ndx+3, 1)))
             ret = ret & chr(((first * 4) and 255) + ((second \ 16) and 3))
             ret = ret & chr(((second * 16) and 255) + ((third \ 4) and 15))
             ret = ret & chr(((third * 64) and 255) + (fourth and 63))
             ndx = ndx + 4
        loop
        if ndx < realLen then
             first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
             second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
             ret = ret & chr(((first * 4) and 255) + ((second \ 16) and 3))
             if realLen mod 4 = 3 then
                  third = Base64DecMap(asc(mid(scrambled,ndx+2,1)))
                  ret = ret & chr(((second * 16) and 255) + ((third \ 4) and 15))
             end if
        end if
        base64Decode = ret
    end function
    call initCodecs
 
    splitvar = split(request("signed_request"), ".")
    encoded_sig = splitvar(0)
    playload = splitvar(1)
     
    set user = JSON.parse(base64Decode(playload))
    session("signed_request") = request("signed_request")
    session("oauth_token") = user.oauth_token
    session("id_atual") = user.user_id
    response.redirect("animacao.asp")
end if
%>