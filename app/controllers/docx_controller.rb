include DocxHelper

class DocxController < ApplicationController
  #unloadable

  def template_upload
    use_for_all = params[:use_for_all]
    uploaded_io = params[:template]
    filename = params[:tracker] + '.docx'
    upload_file(filename, uploaded_io)

    if use_for_all == '1'
      source = Rails.root.join('files', 'export_docx', 'templates', filename)
      Tracker.all.each do |tracker|
        dest = Rails.root.join('files', 'export_docx', 'templates', tracker.name + '.docx')
        FileUtils.copy_file(source, dest) unless source == dest
      end
    end

    redirect_to plugin_settings_path(Redmine::Plugin.find('export_docx'))
  end
  
  def template_download
    tracker = params[:tracker]
    path_to_file = 'files/export_docx/templates/' + tracker + '.docx'
    if File.exist?(path_to_file)
      send_file(path_to_file)
    else
      flash[:error] = 'There is no template for ' + issue.tracker.name + ' issues.'
      redirect_to plugin_settings_path(Redmine::Plugin.find('export_docx'))
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

  private

    def upload_file(filename, uploaded_io)
      if File.extname(uploaded_io.original_filename) == '.docx'
        folder_structure
        File.open(Rails.root.join('files', 'export_docx', 'templates', filename), 'wb') do |file|
          file.write(uploaded_io.read)
          flash[:notice] = filename + ' uploaded successfully.'
        end
      else
        flash[:error] = 'Template must be a .docx file.'
      end
    end
  
end