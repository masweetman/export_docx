require 'docx'

module DocxHelper
  include IssuesHelper
  include CustomFieldsHelper

  def reset_template_names
    templates = list_templates
    templates.each do |t|
      unless t.gsub('.docx', '').last.to_i > 0
        source = t
        dest = t.gsub('.docx', '1.docx')
        FileUtils.copy_file(source, dest)
        FileUtils.rm source
      end
    end
  end

  def upload_file(tracker, uploaded_io, append_template)
    if append_template
      filename = tracker.name + (list_templates_for(tracker).count + 1).to_s + '.docx'
    else
      filename = tracker.name + '1.docx'
    end

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

  def remove_templates_for(tracker)
    list_templates_for(tracker).each do |t|
      FileUtils.rm t
    end
  end

  def list_templates
    Dir.glob 'files/export_docx/templates/*.docx'
  end

  def list_templates_for(tracker)
    templates = []
    list_templates.each do |template|
      filename = template.gsub('files/export_docx/templates/', '')
      templates << template if filename[0..tracker.name.length - 1] == tracker.name && (filename[tracker.name.length].to_i > 0 || filename[tracker.name.length] == '.')
    end
    return templates
  end

  def folder_structure
    FileUtils.mkdir_p 'files/export_docx/templates'
    FileUtils.mkdir_p 'files/export_docx/export'
  end

  def issue_to_docx(issue, template_path)
    FileUtils.rm_rf(Dir.glob('files/export_docx/export/*'))
    path_to_file = template_path.to_s
    if File.exist?(path_to_file)
      doc = Docx::Document.open(path_to_file)
      doc.bookmarks.keys.each do |bookmark|
        # write standard issue fields
        case bookmark.downcase
        when 'project'
          doc.bookmarks[bookmark].insert_text_after(issue.project.name) unless issue.project.nil?
        when 'tracker'
          doc.bookmarks[bookmark].insert_text_after(issue.tracker.name) unless issue.tracker.nil?
        when 'id'
          doc.bookmarks[bookmark].insert_text_after(issue.id.to_s) unless issue.id.nil?
        when 'subject'
          doc.bookmarks[bookmark].insert_text_after(issue.subject) unless issue.subject.nil?
        when 'description'
          doc.bookmarks[bookmark].insert_multiple_lines(issue.description.lines.map(&:chomp)) unless issue.description.nil?
        when 'notes'
          lines = []
          index = 1
          issue.journals.each.with_index do |journal, i|
            lines << "##{index} - #{format_time(journal.created_on)} - #{journal.user}"
            journal.details.each do |detail|
              lines << "- " + show_detail(detail, true)
            end
            lines += journal.notes.lines.map(&:chomp) if journal.notes
            lines << "" if i < issue.journals.size - 1
            index += 1
          end
          doc.bookmarks[bookmark].insert_multiple_lines(lines)
        when 'status'
          doc.bookmarks[bookmark].insert_text_after(issue.status.name) unless issue.status.nil?
        when 'priority'
          doc.bookmarks[bookmark].insert_text_after(issue.priority.name) unless issue.priority.nil?
        when 'author', 'added_by'
          doc.bookmarks[bookmark].insert_text_after(issue.author.name) unless issue.author.nil?
        when 'assignee', 'assigned_to'
          doc.bookmarks[bookmark].insert_text_after(issue.assigned_to.name) unless issue.assigned_to.nil?
        when 'category'
          doc.bookmarks[bookmark].insert_text_after(issue.category.name) unless issue.category.nil?
        when 'target_version', 'fixed_version'
          doc.bookmarks[bookmark].insert_text_after(issue.fixed_version.name) unless issue.fixed_version.nil?
        when 'start_date'
          doc.bookmarks[bookmark].insert_text_after(issue.start_date.strftime('%m/%d/%Y')) unless issue.start_date.nil?
        when 'due_date'
          doc.bookmarks[bookmark].insert_text_after(issue.due_date.strftime('%m/%d/%Y')) unless issue.due_date.nil?
        when 'created_on'
          doc.bookmarks[bookmark].insert_text_after(issue.created_on.strftime('%m/%d/%Y')) unless issue.created_on.nil?
        when 'closed_on'
          doc.bookmarks[bookmark].insert_text_after(issue.closed_on.strftime('%m/%d/%Y')) unless issue.closed_on.nil?
        when 'percent_done', 'done_ratio'
          doc.bookmarks[bookmark].insert_text_after(issue.done_ratio.to_s + '%') unless issue.done_ratio.nil?
        when 'estimated_time', 'estimated_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.estimated_hours.to_s) unless issue.estimated_hours.nil?
        when 'spent_time', 'spent_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.spent_hours.to_s) unless issue.spent_hours.nil?
        else
          #write custom issue fields
          custom_field = CustomField.find_by_name(bookmark.tr('_',' '))
          unless custom_field.nil? || issue.custom_field_value(custom_field.id).nil? || issue.custom_field_value(custom_field.id).empty? || issue.custom_field_value(custom_field.id).first.empty?
            if custom_field.field_format == 'text'
              doc.bookmarks[bookmark].insert_multiple_lines(issue.custom_field_value(custom_field.id).lines.map(&:chomp))
            elsif custom_field.field_format == 'list' && custom_field.multiple?
              doc.bookmarks[bookmark].insert_multiple_lines(issue.custom_field_value(custom_field.id))
            elsif custom_field.field_format == 'user'
              if custom_field.multiple?
                users = []
                users = issue.custom_field_value(custom_field.id).map{ |u| User.find(u).name }
                doc.bookmarks[bookmark].insert_multiple_lines(users) unless users == []
              else
                doc.bookmarks[bookmark].insert_text_after(User.find(issue.custom_field_value(custom_field.id)).to_s)
              end
            elsif custom_field.field_format == 'version'
              if custom_field.multiple?
                versions = []
                versions = issue.custom_field_value(custom_field.id).map{ |v| Version.find(v).name }
                doc.bookmarks[bookmark].insert_multiple_lines(versions) unless versions == []
              else
                doc.bookmarks[bookmark].insert_text_after(Version.find(issue.custom_field_value(custom_field.id)).to_s)
              end
            elsif custom_field.field_format == 'date'
              doc.bookmarks[bookmark].insert_text_after(issue.custom_field_value(custom_field.id).to_date.strftime('%m/%d/%Y'))
            elsif custom_field.field_format == 'bool'
              if doc.bookmarks[bookmark].get_run_before.node.xpath('descendant::*').last.attributes['val'].nil?
                if issue.custom_field_value(custom_field.id) == '1'
                  doc.bookmarks[bookmark].insert_text_after('Yes')
                else
                  doc.bookmarks[bookmark].insert_text_after('No')
                end
              else
                doc.bookmarks[bookmark].get_run_before.node.xpath('descendant::*').last.attributes['val'].value = issue.custom_field_value(custom_field.id).to_s
              end
            else
              doc.bookmarks[bookmark].insert_text_after(issue.custom_field_value(custom_field.id).to_s)
            end
          end
        end
      end
      doc.save('files/export_docx/export/' + issue.project.name + ' - ' + issue.tracker.name + ' #' + issue.id.to_s + '.docx')
    else
      flash[:error] = 'There is no template for ' + issue.tracker.name + ' issues. Please notify your Redmine administrator.'
      redirect_to issue
    end
  end

end