#Access to Facebook using ASP

We need two key files:

Function File access to Facebook (found on the net, do not remember the author, most did not take the header)

Configuration file JSON for ASP (taken from the official site JSON)

You will need a base64 function, I added the Encode and Decode should they need in other projects.

You need to add the ID of your application, the return URL (no need to change the query that put facilitates authentication in a single file), and the permissions in sequence (that you find in the facebook developer).

I saved everything in Session, the token, the signed request and the facebook user ID:


```asp
set user = JSON.parse(base64Decode(playload))
session("signed_request") = request("signed_request")
session("oauth_token") = user.oauth_token
session("id_atual") = user.user_id
response.redirect("pagina.asp")
```
