# Plugin's routes
# See: http://guides.rubyonrails.org/routing.html

get '/issues/:id/export_docx', :to => 'docx#issue_export_docx'
post '/settings/plugin/export_docx/upload', :to => 'docx#template_upload'
get '/settings/plugin/export_docx/download/:tracker', :to => 'docx#template_download'