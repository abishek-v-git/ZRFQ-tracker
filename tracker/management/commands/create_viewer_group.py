from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group


class Command(BaseCommand):
    help = 'Create the Viewer group (read-only role)'

    def handle(self, *args, **options):
        group, created = Group.objects.get_or_create(name='Viewer')
        if created:
            self.stdout.write(self.style.SUCCESS('Viewer group created.'))
        else:
            self.stdout.write('Viewer group already exists.')
        self.stdout.write(
            'Assign users to the Viewer group in Django Admin → '
            'Authentication and Authorization → Users → choose a user → Groups.'
        )
