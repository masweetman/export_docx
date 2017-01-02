Redmine::Plugin.register :export_docx do
  name 'Export DOCX'
  author 'Mike Sweetman'
  description 'This plugin exports issues to DOCX files'
  version '1.1.0'
  url 'http://github.com/masweetman/export_docx'
  author_url 'http://github.com/masweetman'
  
  settings :default => {'empty' => true}, :partial => 'settings/export_docx_settings'
end
