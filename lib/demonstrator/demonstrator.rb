# frozen_string_literal: true
FILENAME_REGEX = /\[(.+?\.xls\|attachment\])\(upload:\/\/(.+?\.xls)\)/

class Demonstrator

  require 'spreadsheet'

  def self.process_topic(topic)
    @process_log = ""
    if can_process_topic
      filename = get_demonstrator_filename(topic)
      demos = get_demonstrator_ids(filename)
      invited_by = User.find(topic.user_id)
      @process_log = ""
      invite_missing(demos, invited_by)
      remove_missing_id(demos)

      notify_complete(topic)
    else
      Rails.logger.error("Something is broken")
    end
  end

  def self.can_process_topic
    demonstrator_group = Group.find_by_name(SiteSetting.demonstrator_group)
    demonstrator_manager_group = Group.find_by_name(SiteSetting.demonstrator_manager_group)
    demonstrator_removed_group = Group.find_by_name(SiteSetting.demonstrator_removed_group)
    demonstrator_category = Category.find(SiteSetting.demonstrator_category)

    (demonstrator_group &&
    demonstrator_manager_group &&
    demonstrator_removed_group &&
    demonstrator_category)
  end

  def self.get_demonstrator_filename(topic)
    local_store = FileStore::LocalStore.new
    post = topic.first_post
    m = FILENAME_REGEX.match(post.raw)
    short_url = m[2]
    u = Upload.find_by(sha1: Upload.sha1_from_short_url(short_url))
    filename = local_store.path_for(u)
  end

  def self.get_demonstrator_ids(filename)
    demos = []
    book = Spreadsheet.open(filename)
    sheet = book.worksheet 0
    email_column = sheet.first.find_index('Email')
    id_column = sheet.first.find_index('Demonstrator ID')
    group_member_column = sheet.first.find_index('Provisionsebene')
    sheet.each 1 do |row|
      email = row[email_column].value if row[email_column].class == Spreadsheet::Formula
      id = row[id_column].value if row[id_column].class == Spreadsheet::Formula
      add_to_group = row[group_member_column].value if row[group_member_column].class == Spreadsheet::Formula

      demos.append({ id: id || row[id_column], email: email || row[email_column], add_to_group: (add_to_group || row[group_member_column]) == 1 })
    end

    demos
  end

  def self.invite_missing(demos, invited_by)
    group = Group.find_by_name(SiteSetting.demonstrator_group)
    @process_log += "## Neue User importieren:\n\n"
    demos.each.with_index(1) do |demo, index|
      next unless demo[:id]
      @process_log += "#{index} (#{demo[:id]}) -> "
      exists_ucf = UserCustomField.find_by(value: demo[:id], name: SiteSetting.demonstrator_ucf)
      if exists_ucf
        @process_log += "Demo-ID #{demo[:id]} existiert schon.\n"
        next
      end

      exists_email = UserEmail.find_by(email: demo[:email].downcase)

      if exists_email
        @process_log += "E-Mail #{demo[:email]} existiert schon.\n"
        next
      end

      invite = Invite.find_by(email: (demo[:email]).downcase)
      if invite
        @process_log += "Einladung an #{demo[:email]} existiert schon.\n"
        next
      end
      opts = {}
      opts[:email] = demo[:email]
      if demo[:add_to_group]
        opts[:group_ids] = [group.id]
      end
      Invite.generate(invited_by, opts)
      @process_log += "\n**Eingeladen** #{demo[:email]}\n\n "
    end
    @process_log += "\nKeine weiteren neuen User\n\n\n"
  end

  def self.remove_missing_id(demos)
    @process_log += "## User entfernen\n\n"
    demonstrator_ids = demos.map { |i| (i[:id]).to_i }
    manager_group = Group.find_by_name(SiteSetting.demonstrator_manager_group)
    removed_group = Group.find_by_name(SiteSetting.demonstrator_removed_group)
    demo_group = Group.find_by_name(SiteSetting.demonstrator_group)
    users = User.all
    users.each do |user|
      next if user.staff?
      next if GroupUser.find_by(user_id: user.id, group_id: manager_group.id)
      ucf = UserCustomField.find_by(user_id: user.id, name: SiteSetting.demonstrator_ucf)
      if ucf
        gu = GroupUser.find_by(user_id: user.id, group_id: demo_group.id)
        id = (demos.select { |i| i['id'] == ucf.value }).first
        if id
          if id['add_to_group']
            GroupUser.find_or_create_by(user_id: user.id, group_id: demo_group.id)
            @process_log += ("#{user.username} => #{demo_group.name}\n")
          else
            gu.destroy if gu
            @process_log += ("#{user.username} XXX #{demo_group.name}\n")
          end
        end
      end
      next if ucf && demonstrator_ids.include?(ucf.value.to_i)
      ucf.value = ""
      ucf.save
      user.email = "#{user.username}@removed.invalid"
      user.active = false
      user.save
      demo_group_user = GroupUser.find_by(user_id: user.id, group_id: demo_group.id)
      demo_group_user.destroy if demo_group_user
      GroupUser.find_or_create_by(user_id: user.id, group_id: removed_group.id)
      @process_log += "#{user.username}\n"
    end
  end

  def self.notify_complete(topic)
    post = PostCreator.new(
      User.find(-1),
      raw: @process_log,
      skip_jobs: true,
      skip_validations: true,
      topic_id: topic.id,
    ).create
  end
end
