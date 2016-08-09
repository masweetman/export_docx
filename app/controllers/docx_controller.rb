include DocxHelper

class DocxController < ApplicationController
  #unloadable

  def template_upload
    uploaded_io = params[:template]
    filename = params[:tracker] + '.docx'
    if File.extname(uploaded_io.original_filename) == '.docx'
      folder_structure
      File.open(Rails.root.join('files', 'export_docx', 'templates', filename), 'wb') do |file|
        file.write(uploaded_io.read)
        flash[:notice] = filename + ' uploaded successfully.'
      end
    else
      flash[:error] = 'Template must be a .docx file.'
    end
    redirect_to '/settings/plugin/export_docx'
  end
  
  def template_download
    tracker = params[:tracker]
    path_to_file = 'files/export_docx/templates/' + tracker + '.docx'
    if File.exist?(path_to_file)
      send_file(path_to_file)
    else
      flash[:error] = 'A template for ' + tracker + ' issues does not exist.'
      redirect_to '/settings/plugin/export_docx'
    end
  end
  
  def issue_export_docx
    issue = Issue.find(params[:id])
    issue_to_docx(issue)
    path_to_file = 'files/export_docx/export/' + issue.project.name + ' - ' + issue.tracker.name + ' #' + issue.id.to_s + '.docx'
    if File.exist?(path_to_file)
      send_file(path_to_file)
    end
  end
  
end