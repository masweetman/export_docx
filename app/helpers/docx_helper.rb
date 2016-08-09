require 'docx'

module DocxHelper

  def issue_to_docx(issue)
    FileUtils.rm_rf(Dir.glob('files/export_docx/export/*'))
    path_to_file = 'files/export_docx/templates/' + issue.tracker.name + '.docx'
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
        when 'status'
          doc.bookmarks[bookmark].insert_text_after(issue.status.name) unless issue.status.nil?
        when 'priority'
          doc.bookmarks[bookmark].insert_text_after(issue.priority.name) unless issue.priority.nil?
        when 'author', 'added by', 'added_by'
          doc.bookmarks[bookmark].insert_text_after(issue.author.name) unless issue.author.nil?
        when 'assignee', 'assigned to', 'assigned_to'
          doc.bookmarks[bookmark].insert_text_after(issue.assigned_to.name) unless issue.assigned_to.nil?
        when 'category'
          doc.bookmarks[bookmark].insert_text_after(issue.category.name) unless issue.category.nil?
        when 'target version', 'target_version', 'fixed version' 'fixed_version'
          doc.bookmarks[bookmark].insert_text_after(issue.fixed_version.name) unless issue.fixed_version.nil?
        when 'start date', 'start_date'
          doc.bookmarks[bookmark].insert_text_after(issue.start_date.strftime('%m/%d/%Y')) unless issue.start_date.nil?
        when 'due date', 'due_date'
          doc.bookmarks[bookmark].insert_text_after(issue.due_date.strftime('%m/%d/%Y')) unless issue.due_date.nil?
        when '% done', '%_done', 'percent done', 'percent_done', 'done ratio', 'done_ratio'
          doc.bookmarks[bookmark].insert_text_after(issue.done_ratio.to_s + '%') unless issue.done_ratio.nil?
        when 'estimated time', 'estimated_time', 'estimated hours', 'estimated_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.estimated_hours.to_s) unless issue.estimated_hours.nil?
        when 'spent time', 'spent_time', 'spent hours', 'spent_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.spent_hours.to_s) unless issue.spent_hours.nil?
        else
          #write custom issue fields
          custom_field = CustomField.find_by_name(bookmark)
          unless custom_field.nil?
            if custom_field.field_format == 'text'
              doc.bookmarks[bookmark].insert_multiple_lines(issue.custom_field_value(custom_field.id).lines.map(&:chomp))
            else
              doc.bookmarks[bookmark].insert_text_after(issue.custom_field_value(custom_field.id).to_s)
            end
          end
        end
      end
      doc.save('files/export_docx/export/'+ issue.project.name + ' - ' + issue.tracker.name + ' #' + issue.id.to_s + '.docx')
    else
      flash[:error] = 'A template for ' + issue.tracker.name + ' issues does not exist. Please notify your Redmine administrator.'
      redirect_to issue
    end
  end

end